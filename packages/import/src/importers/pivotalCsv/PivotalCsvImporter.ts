import csv from "csvtojson";
import { Importer, ImportResult } from "../../types";

type PivotalStoryType = "epic" | "feature" | "bug" | "chore" | "release";

interface PivotalIssueType {
  Id: string;
  Title: string;
  Labels: string;
  Iteration: string;
  "Iteration Start": string;
  "Iteration End": string;
  Type: PivotalStoryType;
  Estimate: string;
  "Current State": string;
  "Created at": Date;
  "Accepted at": Date;
  Deadline: string;
  "Requested By": string;
  Description: string;
  URL: string;
  "Owned By": string;
  Blocker: string;
  "Blocker Status": string;
  Comment: string;
}

// Pivotal Tracker CSV columns:
// - Id [integer]
// - Title [string]
// - Labels [string] -- CSV
// - Iteration [integer]
// - Iteration Start [date]
// - Iteration End [date]
// - Type [string] -- e.g. "feature"
// - Estimate [integer]
// - Priority [string] -- e.g. "p3 - Low"
// - Current State [string]
// - Created at [date]
// - Accepted at [date]
// - Deadline [date]
// - Requested By [string] -- user name
// - Description [string]
// - URL [string]
// - Owned By [string] -- user name
// - Blocker [string] [repeated column]
// - Blocker Status [string] [repeated column] -- "resolved" or blank
// - Comment [string] [repeated column]
// - Task [string] [repeated column]
// - Task Status [string] [repeated column] -- "completed" or "not completed"
// - Review Type [string] [repeated column]
// - Reviewer [string] [repeated column] -- user name
// - Review Status [string] [repeated column] --"pass", "unstarted", "revise", or "in_review"
// - Pull Request [string] [repeated column] -- URL
// - Git Branch [string] [repeated column]

// Handle duplicate column headers, which will be renamed like "Column 1", "Column 2".
// Those renamed headers are tracked here.
const duplicateHeaders = {
  Blocker: [],
  Comment: [],
  "Owned By": [],
  "Pull Request": [],
  Task: [],
};

/**
 * Import issues from an Pivotal Tracker CSV export.
 *
 * @param filePath  path to csv file
 * @param orgSlug   base Pivotal project url
 */
export class PivotalCsvImporter implements Importer {
  public constructor(filePath: string) {
    this.filePath = filePath;
  }

  public get name(): string {
    return "Pivotal (CSV)";
  }

  public get defaultTeamName(): string {
    return "Pivotal";
  }

  public import = async (): Promise<ImportResult> => {
    const csvPromise = csv()
      .fromFile(this.filePath)
      // Rename and track the duplicate headers, e.g. "Comment" => "Comment 1", "Comment 2"
      .on("header", headers => {
        headers.forEach((header: string, i: number) => {
          const aliases = duplicateHeaders[header];
          if (aliases) {
            const alias = `${header} ${aliases.length + 1}`;
            headers[i] = alias;
            aliases.push(alias);
          }
        });
      });

    const data = (await csvPromise) as PivotalIssueType[];

    const importData: ImportResult = {
      issues: [],
      labels: {},
      users: {},
    };

    const assignees = Array.from(new Set(data.map(row => row["Owned By 1"])));

    for (const user of assignees) {
      importData.users[user] = {
        name: user,
      };
    }

    for (const row of data) {
      const type = row.Type;
      if (type === "epic" || type === "release") {
        continue;
      }

      let title = row.Title;
      if (!title) {
        continue;
      }

      const id = row["Id"];

      // Add the PT story ID as a suffix to the title
      title += ` [PT #${id}]`;

      let description = row.Description;

      const url = row.URL;

      const originalEstimate = parseInt(row["Estimate"]);
      const estimate = mapEstimate(originalEstimate);

      const priority = mapPriority(row["Priority"]);

      const labels = row.Labels.split(",")
        .map(s => s.trim())
        .filter(Boolean);

      const creatorId = row["Requested By"] || undefined;

      // Handle only the first "Owned By" column (which we rename to "Owned By 1"
      // in the headers processing)
      const assigneeId = mapAssigneeId(row["Owned By 1"]) || undefined;

      const status = mapStatus(row["Current State"], labels);

      // add the type as a label
      labels.push(type);

      const createdAt = row["Created at"];
      const completedAt = row["Accepted at"];

      // Add PT values to the start of the description
      const ptDescItems = [];
      if (url) {
        ptDescItems.push(`PT [#${id}](${url})`);
      }
      if (originalEstimate) {
        ptDescItems.push(`PT estimate: ${originalEstimate}`);
      }
      if (creatorId) {
        ptDescItems.push(`PT creator: ${creatorId}`);
      }
      if (ptDescItems.length > 0) {
        description = `${ptDescItems.join("\n")}\n\n${description}`;
      }

      // Extract tasks, blockers, and PR URLs, and add them to the end of the description
      const tasks = duplicateHeaderValues("Task", row)
        .map((task: string) => `\n- ${task}`)
        .join("");
      if (tasks) {
        description += `\n\n**Tasks**:${tasks}`;
      }

      const blockers = duplicateHeaderValues("Blocker", row)
        .map((blocker: string) => `\n- ${blocker}`)
        .join("");
      if (blockers) {
        description += `\n\n**Blockers**:${blockers}`;
      }

      const prUrls = duplicateHeaderValues("Pull Request", row)
        .map((prUrl: string) => `\n- ${prUrl}`)
        .join("");
      if (prUrls) {
        description += `\n\n**Pull Requests**:${prUrls}`;
      }

      // Extract and transform the comments.
      //
      // The exported comments include the author name and creation date
      // appended, e.g. "(Alice Smith - Apr 1, 2024)"
      const commentMetaRegex = /\s*\(([\w\s]+) - (\w{3} \d+, \d\d\d\d)\)$/;
      const comments = duplicateHeaderValues("Comment", row).map((comment: string) => {
        const matches = comment.match(commentMetaRegex) || [];
        const userId = matches[1];
        let commentCreatedAt = new Date(matches[2]);
        if (isNaN(commentCreatedAt.getTime()) || commentCreatedAt > new Date()) {
          commentCreatedAt = new Date();
        }

        comment = comment.replace(commentMetaRegex, "");

        return {
          body: comment,
          createdAt: commentCreatedAt,
          userId,
        };
      });

      importData.issues.push({
        title,
        description,
        estimate,
        priority,
        status,
        url,
        assigneeId,
        labels,
        createdAt,
        completedAt,
        comments,
      });

      for (const lab of labels) {
        if (!importData.labels[lab]) {
          importData.labels[lab] = {
            name: lab,
          };
        }
      }
    }

    return importData;
  };

  // -- Private interface

  private filePath: string;
}

const duplicateHeaderValues = (header: string, row: PivotalIssueType) => {
  return duplicateHeaders[header].map((alias: string) => row[alias]).filter(Boolean);
};

const mapAssigneeId = (input: string) => {
  const userMap = {
    aaron1V: "aaron@frontporchforum.com",
    "Ariel Wish": "ariel@frontporchforum.com",
    "Brendan McKay": "brendan@frontporchforum.com",
    "Bridget Mientka": "bridget@frontporchforum.com",
    "Chloe Tomlinson": "chloe@frontporchforum.com",
    danielle1I: "danielle@frontporchforum.com",
    Emily: "emily@frontporchforum.com",
    "eva e": "eva@frontporchforum.com",
    gisele5: "gisele@frontporchforum.com",
    james3F: "james@frontporchforum.com",
    jan26: "jan@frontporchforum.com",
    "Jason Van Driesche": "jason@frontporchforum.com",
    jillX: "jill@frontporchforum.com",
    jonnajermyn: "jonna.jermyn@frontporchforum.com",
    "Kathryn Goulding": "Kathryn@frontporchforum.com",
    martaQ: "marta@frontporchforum.com",
    Matt: "Matt Barry",
    michaelfpf: "michael@frontporchforum.com",
    Nina: "nina@frontporchforum.com",
    "Noah Harrison": "noah@frontporchforum.com",
    stefan: "stefan@frontporchforum.com",
    susannah1: "susannah@frontporchforum.com",
    wendy18: "wendy@frontporchforum.com",
  };

  return userMap[input] || input;
};

const mapEstimate = (input: number) => {
  // Values > 64 result in a validation error
  return input && input <= 64 ? input : undefined;
};

const mapPriority = (input: string) => {
  const priorityMap = {
    none: 0,
    "p0 - Critical": 1,
    "p1 - High": 2,
    "p2 - Medium": 3,
    "p3 - Low": 4,
  };
  return priorityMap[input] || 0;
};

const mapStatus = (input: string, labels: string[]) => {
  if (input === "accepted" && labels.includes("abandoned")) {
    return "Canceled";
  }

  if (input === "accepted" && labels.includes("released")) {
    return "Released";
  }

  const statusMap = {
    unscheduled: "Icebox",
    unstarted: "Backlog",
    planned: "Todo",
    started: "In Progress",
    finished: "Dev Review",
    delivered: "QA Review",
    accepted: "Accepted",
  };
  return statusMap[input] || "Backlog";
};
