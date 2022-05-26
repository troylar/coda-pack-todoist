import * as coda from "@codahq/packs-sdk";
import { v4 as uuidv4 } from 'uuid';

// Constants.

const ProjectUrlPatterns: RegExp[] = [
    new RegExp("^https://todoist.com/app/project/([0-9]+)$"),
    new RegExp("^https://todoist.com/showProject\\?id=([0-9]+)"),
];

const TaskUrlPatterns: RegExp[] = [
    new RegExp("^https://todoist.com/app/project/[0-9]+/task/([0-9]+)$"),
    new RegExp("^https://todoist.com/showTask\\?id=([0-9]+)"),
];

// Pack setup.

export const pack = coda.newPack();

pack.addNetworkDomain("todoist.com");

pack.setUserAuthentication({
    type: coda.AuthenticationType.OAuth2,
    // OAuth2 URLs and scopes are found in the the Todoist OAuth guide:
    // https://developer.todoist.com/guides/#oauth
    authorizationUrl: "https://todoist.com/oauth/authorize",
    tokenUrl: "https://todoist.com/oauth/access_token",
    scopes: ["data:read_write"],
    scopeDelimiter: ",",

    // Determines the display name of the connected account.
    getConnectionName: async function (context) {
        let url = coda.withQueryParams("https://api.todoist.com/sync/v8/sync", {
            resource_types: JSON.stringify(["user"]),
        });
        let response = await context.fetcher.fetch({
            method: "GET",
            url: url,
        });
        return response.body.user?.full_name;
    },
});

const ActivityLogEventSchema = coda.makeObjectSchema({
    properties: {
        id: {
            description: "The ID of the event.",
            type: coda.ValueType.Number,
        },
        name: {
            description: "Event object name.",
            type: coda.ValueType.String,
        },
        note: {
            description: "Note content.",
            type: coda.ValueType.String,
        },
        objectType: {
            description: "The type of object, one of item, note or project.",
            type: coda.ValueType.String,
        },
        objectId: {
            description: "The ID of the object.",
            type: coda.ValueType.Number,
        },
        eventType: {
            description: "The type of event, one of added, updated, deleted, completed, uncompleted, archived, unarchived, shared, left.",
            type: coda.ValueType.String,
        },
        eventDate: {
            description: "The date and time when the event took place.",
            type: coda.ValueType.String,
        },
        parentProjectId: {
            description: "The ID of the item's or note's parent project, otherwise null.",
            type: coda.ValueType.Number,
        },
        parentItemId: {
            description: "The ID of the note's parent item, otherwise null.",
            type: coda.ValueType.Number,
        },
        initiatorId: {
            description: "The ID of the user who is responsible for the event, which only makes sense in shared projects, items and notes, and is null for non-shared objects.",
            type: coda.ValueType.Number,
        },
        extraData: {
            description: "This object contains at least the name of the project, or the content of an item or note, and optionally the last_name if a project was renamed, the last_content if an item or note was renamed, the due_date and last_due_date if an item's due date changed, the responsible_uid and last_responsible_uid if an item's responsible uid changed, the description and last_description if an item's description changed, and the client that caused the logging of the event.",
            type: coda.ValueType.Number,
        },
        url: {
            description: "Object URL",
            type: coda.ValueType.String,
        },
    },
    displayProperty: "objectType",
    idProperty: "id",
    identity: {
        name: "ActivityLogEvent",
    }
})

/**
 * Convert a ActivityLog API response to a ActivityLogEvent schema.
 */
function toActivityLogEvent(event: any) {
    let result: any = {
        id: event.id,
        objectType: event.object_type,
        objectId: event.object_id,
        eventType: event.event_type,
        eventDate: event.event_date,
        parentProjectId: event.parent_project_id,
        parentItemId: event.parent_item_id,
        initiatorId: event.initiator_id,
        extraData: event.extra_data,
        url: event.url
    };

    if (event.extra_data.hasOwnProperty('name')) {
        result.name = event.extra_data.name;
    }
    if (event.extra_data.hasOwnProperty('content')) {
        result.note = event.extra_data.content;
    }
    return result;
}

pack.addSyncTable({
    name: "ActivityLog",
    schema: ActivityLogEventSchema,
    identityName: "ActivityLogEvent",
    formula: {
        name: "SyncActivityLog",
        description: "Sync Activity Log",
        parameters: [],
        execute: async function ([], context) {
            let limit = 100;
            let offset = 0;
            let done = false;
            let results = [];
            while (!done) {
                console.log('Getting activity offset=' + offset + ', limit=' + limit);
                let url = "https://api.todoist.com/sync/v8/activity/get?sync_token=*&offset=" + offset + "&limit=" + limit;
                console.log(url);
                let response = await context.fetcher.fetch({
                    method: "GET",
                    url: url
                });
                let object: any;
                response.body.events.forEach(async function (event) {
                    if (event.object_type == "item") {
                        event.url = "https://todoist.com/app/task/" + event.object_id
                    }
                    if (event.object_type == "note") {
                        event.url = "https://todoist.com/app/today/task/" + event.parent_item_id + "/comments#comment-" + event.object_id;
                    }
                    if (event.object_type == "project") {
                        event.url = "https://todoist.com/app/project/" + event.object_id;
                    }
                    results.push(toActivityLogEvent(event));
                });
                if (response.body.events.length < limit) {
                    done = true;
                }
                else {
                    offset += limit;
                }
            }
            return {
                result: results,
            };
        },
    },
});


const SharedLabelSchema = coda.makeObjectSchema({
    properties: {
        name: {
            description: "The name of the label.",
            type: coda.ValueType.String,
            required: true,
        },
    },
    displayProperty: "name",
    idProperty: "name",
    identity: {
        name: "SharedLabel",
    }
});
// Schemas

const LabelSchema = coda.makeObjectSchema({
    properties: {
        name: {
            description: "The name of the label.",
            type: coda.ValueType.String,
            required: true,
        },
        labelId: {
            description: "The ID of the label.",
            type: coda.ValueType.Number,
            required: true,
        },
        color: {
            description: "The color of the label.",
            type: coda.ValueType.Number,
            required: true,
        },
        order: {
            description: "The order of the label.",
            type: coda.ValueType.Number,
            required: true,
        },
        favorite: {
            description: "Is this a favorite?",
            type: coda.ValueType.Boolean,
            required: true,
        }
    },
    displayProperty: "name",
    idProperty: "labelId",
    identity: {
        name: "Label",
    }
});
// Schemas

// A reference to a synced Project. Usually you can use
// `coda.makeReferenceSchemaFromObjectSchema` to generate these from the primary
// schema, but that doesn't work in this case since a Project itself can contain
// a reference to a parent project.
const ProjectReferenceSchema = coda.makeObjectSchema({
    codaType: coda.ValueHintType.Reference,
    properties: {
        name: { type: coda.ValueType.String, required: true },
        projectId: { type: coda.ValueType.Number, required: true },
    },
    displayProperty: "name",
    idProperty: "projectId",
    identity: {
        name: "Project",
    },
});

const ProjectSchema = coda.makeObjectSchema({
    properties: {
        name: {
            description: "The name of the project.",
            type: coda.ValueType.String,
            required: true,
        },
        url: {
            description: "A link to the project in the Todoist app.",
            type: coda.ValueType.String,
            codaType: coda.ValueHintType.Url,
        },
        shared: {
            description: "Is the project is shared.",
            type: coda.ValueType.Boolean,
        },
        favorite: {
            description: "Is the project a favorite.",
            type: coda.ValueType.Boolean,
        },
        projectId: {
            description: "The ID of the project.",
            type: coda.ValueType.Number,
            required: true,
        },
        parentProjectId: {
            description: "For sub-projects, the ID of the parent project.",
            type: coda.ValueType.Number,
        },
        // Add a reference to the sync'ed row of the parent project.
        // References only work in sync tables.
        parentProject: ProjectReferenceSchema,
    },
    displayProperty: "name",
    idProperty: "projectId",
    featuredProperties: ["url"],
    identity: {
        name: "Project",
    },
});

// A reference to a synced Task. Usually you can use
// `coda.makeReferenceSchemaFromObjectSchema` to generate these from the primary
// schema, but that doesn't work in this case since a task itself can contain
// a reference to a parent task.
const TaskReferenceSchema = coda.makeObjectSchema({
    codaType: coda.ValueHintType.Reference,
    properties: {
        name: { type: coda.ValueType.String, required: true },
        taskId: { type: coda.ValueType.Number, required: true },
    },
    displayProperty: "name",
    idProperty: "taskId",
    identity: {
        name: "Task",
    },
});

const TaskSchema = coda.makeObjectSchema({
    properties: {
        name: {
            description: "The name of the task.",
            type: coda.ValueType.String,
            required: true,
        },
        description: {
            description: "A detailed description of the task.",
            type: coda.ValueType.String,
        },
        url: {
            description: "A link to the task in the Todoist app.",
            type: coda.ValueType.String,
            codaType: coda.ValueHintType.Url,
        },
        order: {
            description: "The position of the task in the project or parent task.",
            type: coda.ValueType.Number,
        },
        priority: {
            description: "The priority of the task.",
            type: coda.ValueType.String,
        },
        taskId: {
            description: "The ID of the task.",
            type: coda.ValueType.Number,
            required: true,
        },
        projectId: {
            description: "The ID of the project that the task belongs to.",
            type: coda.ValueType.Number,
        },
        parentTaskId: {
            description: "For sub-tasks, the ID of the parent task it belongs to.",
            type: coda.ValueType.Number,
        },
        checked: {
            description: "Is the task completed?",
            type: coda.ValueType.Boolean,
        },
        assignee: {
            description: "The ID of user who is responsible for accomplishing the current task.",
            type: coda.ValueType.Number,
        },
        labels: {
            type: coda.ValueType.Array,
            items: coda.makeSchema({
                description: "The tasks labels (a list of label IDs such as [2324,2525]).",
                type: coda.ValueType.Number,
            }),
        },
        dateAdded: {
            description: "Creation date",
            type: coda.ValueType.String
        },
        dateCompleted: {
            description: "Completion date (if completed)",
            type: coda.ValueType.String
        },
        due: {
            description: "The due date of the task.",
            type: coda.ValueType.Object,
            properties: {
                date: {
                    description: "Due date in the format of YYYY-MM-DD (RFC 3339). For recurring dates, the date of the current iteration.",
                    type: coda.ValueType.String
                },
                timezone: {
                    description: "Always set to null.",
                    type: coda.ValueType.String
                },
                string: {
                    description: "Human-readable representation of due date. String always represents the due object in user's timezone.",
                    type: coda.ValueType.String
                },
                lang: {
                    description: "Lang which has to be used to parse the content of the string attribute.",
                    type: coda.ValueType.String
                },
                is_recurring: {
                    description: "Boolean flag which is set to true if the due object represents a recurring due date.",
                    type: coda.ValueType.Boolean
                }
            },
        },
        // A reference to the sync'ed row of the project.
        // References only work in sync tables.
        project: ProjectReferenceSchema,
        // Add a reference to the sync'ed row of the parent task.
        // References only work in sync tables.
        parentTask: TaskReferenceSchema,
    },
    displayProperty: "name",
    idProperty: "taskId",
    featuredProperties: ["project", "url"],
    identity: {
        name: "Task",
    },
});

/**
 * Convert a Project API response to a Project schema.
 */
function toLabel(label: any) {
    let result: any = {
        name: label.name,
        labelId: label.id,
        color: label.color,
        order: label.shared,
        favorite: label.favorite,
    };
    return result;
}

/**
 * Convert a Project API response to a Project schema.
 */
function toSharedLabel(label: any) {
    let result: any = {
        name: label.name,
    };
    return result;
}

/**
 * Convert a Project API response to a Project schema.
 */
function toProject(project: any, withReferences = false) {
    let result: any = {
        name: project.name,
        projectId: project.id,
        url: project.url,
        shared: project.shared,
        favorite: project.favorite,
        parentProjectId: project.parent_id,
    };
    if (withReferences && project.parent_id) {
        result.parentProject = {
            projectId: project.parent_id,
            name: "Not found", // If sync'ed, the real name will be shown instead.
        };
    }
    return result;
}

/**
 * Convert a Task API response to a Task schema.
 */
function toTask(task: any, withReferences = false) {
    let result: any = {
        name: task.content,
        description: task.description,
        url: task.url,
        order: task.order,
        priority: task.priority,
        taskId: task.coda_sync_id,
        projectId: task.project_id,
        parentTaskId: task.parent_id,
        labels: task.labels,
        due: task.due,
        checked: task.checked == true,
        dateAdded: task.date_added,
        dateCompleted: task.completed_date
    };
    if (task.responsible_uid) {
        result.assignee = task.responsible_uid
    }

    if (withReferences) {
        // Add a reference to the corresponding row in the Projects sync table.
        result.project = {
            projectId: task.project_id,
            name: "Not found", // If sync'ed, the real name will be shown instead.
        };
        if (task.parent_id) {
            // Add a reference to the corresponding row in the Tasks sync table.
            result.parentTask = {
                taskId: task.parent_id,
                name: "Not found", // If sync'ed, the real name will be shown instead.
            };
        }
    }
    return result;
}

pack.addFormula({
    name: "GetProject",
    description: "Gets a Todoist project by URL",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "url",
            description: "The URL of the project",
        }),
    ],
    resultType: coda.ValueType.Object,
    schema: ProjectSchema,

    execute: async function ([url], context) {
        let projectId = extractProjectId(url);
        let response = await context.fetcher.fetch({
            url: "https://api.todoist.com/rest/v1/projects/" + projectId,
            method: "GET",
        });
        return toProject(response.body);
    },
});

pack.addFormula({
    name: "GetTask",
    description: "Gets a Todoist task by URL",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "url",
            description: "The URL of the task",
        }),
    ],
    resultType: coda.ValueType.Object,
    schema: TaskSchema,

    execute: async function ([url], context) {
        let taskId = extractTaskId(url);
        let response = await context.fetcher.fetch({
            url: "https://api.todoist.com/rest/v1/tasks/" + taskId,
            method: "GET",
        });
        return toTask(response.body);
    },
});

// Column Formats.

pack.addColumnFormat({
    name: "Project",
    formulaName: "GetProject",
    matchers: ProjectUrlPatterns,
});

pack.addColumnFormat({
    name: "Task",
    formulaName: "GetTask",
    matchers: TaskUrlPatterns,
});

// Action formulas (buttons/automations).

pack.addFormula({
    name: "AddProject",
    description: "Add a new Todoist project",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "name",
            description: "The name of the new project",
        }),
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "collaborators",
            description: "Comma-delimited list of collaborator emails",
        }),
    ],
    resultType: coda.ValueType.String,
    isAction: true,
    extraOAuthScopes: ["data:read_write"],

    execute: async function ([name, collaborators], context) {
        let response = await context.fetcher.fetch({
            url: "https://api.todoist.com/rest/v1/projects",
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify({
                name: name,
            }),
        });
        let project_id = response.body.id;
        let project_url = response.body.url;
        collaborators.split(",").forEach(async function (collaborator) {
            console.log("Sharing project with " + collaborator);
            let commands = JSON.stringify([{ type: "share_project", uuid: uuidv4(), args: { project_id: project_id, email: collaborator } }]);
            let url = "https://api.todoist.com/sync/v8/sync?sync_token=*&commands=" + commands;
            response = await context.fetcher.fetch({
                method: "POST",
                url: url,
            });
        });
        return project_url;
    },
});

pack.addFormula({
    name: "AddTask",
    description: "Add a new task.",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "name",
            description: "The name of the task.",
        }),
        coda.makeParameter({
            type: coda.ParameterType.Number,
            name: "projectId",
            description: "The ID of the project to add it to. If blank, " +
                "it will be added to the user's Inbox.",
            optional: true,
        }),
    ],
    resultType: coda.ValueType.String,
    isAction: true,
    extraOAuthScopes: ["data:read_write"],

    execute: async function ([name, projectId], context) {
        let response = await context.fetcher.fetch({
            url: "https://api.todoist.com/rest/v1/tasks",
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify({
                content: name,
                project_id: projectId,
            }),
        });
        return response.body.url;
    },
});

pack.addFormula({
    name: "UpdateTask",
    description: "Updates the name of a task.",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "taskId",
            description: "The ID of the task to update.",
        }),
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "name",
            description: "The new name of the task.",
        }),
    ],
    resultType: coda.ValueType.Object,
    schema: TaskSchema,
    isAction: true,
    extraOAuthScopes: ["data:read_write"],

    execute: async function ([taskId, name], context) {
        let url = "https://api.todoist.com/rest/v1/tasks/" + taskId;
        await context.fetcher.fetch({
            url: url,
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify({
                content: name,
            }),
        });
        // Get the updated Task and return it, which will update the row in the sync
        // table.
        let response = await context.fetcher.fetch({
            url: url,
            method: "GET",
            cacheTtlSecs: 0, // Ensure we are getting the latest data.
        });
        return toTask(response.body);
    },
});

pack.addFormula({
    name: "MarkAsComplete",
    description: "Mark a task as completed.",
    parameters: [
        coda.makeParameter({
            type: coda.ParameterType.String,
            name: "taskId",
            description: "The ID of the task to be marked as complete.",
        }),
    ],
    resultType: coda.ValueType.String,
    isAction: true,
    extraOAuthScopes: ["data:read_write"],

    execute: async function ([taskId], context) {
        let url = "https://api.todoist.com/rest/v1/tasks/" + taskId + "/close";
        await context.fetcher.fetch({
            method: "POST",
            url: url,
            headers: {
                "Content-Type": "application/json",
            },
        });
        return "OK";
    },
});

// Sync tables.
pack.addSyncTable({
    name: "SharedLabels",
    schema: SharedLabelSchema,
    identityName: "SharedLabel",
    formula: {
        name: "SharedLabels",
        description: "Sync shared labels",
        parameters: [],

        execute: async function ([], context) {
            let url = "https://api.todoist.com/sync/v8/sync?sync_token=*&resource_types=[\"items\"]";
            let response = await context.fetcher.fetch({
                method: "GET",
                url: url,
            });

            let results = [];
            for (let item of response.body.items) {
                for (let labelName of item.labels) {
                    const index = results.findIndex(o => o.name === labelName);
                    if (index == -1) {
                        console.log('Adding ' + labelName)
                        results.push(toSharedLabel({ name: labelName }))
                    }
                    else {
                        console.log('Skipping ' + labelName)
                        console.log(results)
                    }
                }
            }
            return {
                result: results,
            };
        },
    },
});

pack.addSyncTable({
    name: "Labels",
    schema: LabelSchema,
    identityName: "Label",
    formula: {
        name: "SyncLabels",
        description: "Sync labels",
        parameters: [],

        execute: async function ([], context) {
            let url = "https://api.todoist.com/sync/v8/sync?sync_token=*&resource_types=[\"labels\",\"items\",\"projects\",\"collaborators\"]";
            let response = await context.fetcher.fetch({
                method: "GET",
                url: url,
            });

            let results = [];
            for (let item of response.body.items) {
                for (let labelName of item.labels) {
                    const index = results.findIndex(o => o.name === labelName);
                    if (index == -1) {
                        console.log('Adding ' + labelName)
                        results.push(toLabel({ name: labelName }))
                    }
                    else {
                        console.log('Skipping ' + labelName)
                    }
                }
            }
            for (let label of response.body.labels) {
                const index = results.findIndex(o => o.name === label.name);
                if (index == -1) {
                    results.push(toLabel(label));
                }
            }
            return {
                result: results,
            };
        },
    },
});

pack.addSyncTable({
    name: "Projects",
    schema: ProjectSchema,
    identityName: "Project",
    formula: {
        name: "SyncProjects",
        description: "Sync projects",
        parameters: [],

        execute: async function ([], context) {
            let url = "https://api.todoist.com/rest/v1/projects";
            let response = await context.fetcher.fetch({
                method: "GET",
                url: url,
            });

            let results = [];
            for (let project of response.body) {
                results.push(toProject(project, true));
            }
            return {
                result: results,
            };
        },
    },
});

pack.addSyncTable({
    name: "Tasks",
    schema: TaskSchema,
    identityName: "Task",
    formula: {
        name: "SyncTasks",
        description: "Sync tasks",
        parameters: [
            coda.makeParameter({
                type: coda.ParameterType.String,
                name: "filter",
                description: "A supported filter string. See the Todoist help center.",
                optional: true,
            }),
            coda.makeParameter({
                type: coda.ParameterType.String,
                name: "project",
                description: "Limit tasks to a specific project.",
                optional: true,
                autocomplete: async function (context, search) {
                    let url = "https://api.todoist.com/rest/v1/projects";
                    let response = await context.fetcher.fetch({
                        method: "GET",
                        url: url,
                    });
                    let projects = response.body;
                    return coda.autocompleteSearchObjects(search, projects, "name", "id");
                },
            }),
        ],
        execute: async function ([filter, project], context) {
            let url = "https://api.todoist.com/sync/v8/sync?sync_token=*&resource_types=[\"items\"]";

            let response = await context.fetcher.fetch({
                method: "GET",
                url: url,
            });

            let results = [];
            for (let task of response.body.items) {
                task.coda_sync_id = task.sync_id ?? task.id;
                results.push(toTask(task, true));
            }

            url = "https://api.todoist.com/sync/v8/completed/get_all";

            response = await context.fetcher.fetch({
                method: "GET",
                url: url,
            });

            for (let task of response.body.items) {
                task.checked = true;
                task.coda_sync_id = task.task_id;
                results.push(toTask(task, true));
            }

            return {
                result: results,
            };
        },
    },
});

// Helper functions.

function extractProjectId(projectUrl: string) {
    for (let pattern of ProjectUrlPatterns) {
        let matches = projectUrl.match(pattern);
        if (matches && matches[1]) {
            return matches[1];
        }
    }
    throw new coda.UserVisibleError("Invalid project URL: " + projectUrl);
}

function extractTaskId(taskUrl: string) {
    for (let pattern of TaskUrlPatterns) {
        let matches = taskUrl.match(pattern);
        if (matches && matches[1]) {
            return matches[1];
        }
    }
    throw new coda.UserVisibleError("Invalid task URL: " + taskUrl);
}