// Brillienta-Helpdesk in the suite agent: the same Graph calls, fields and queued actions as Helpdesk's own agent
// (brillientahelpdeskwebpart/src/webparts/helpdesk/services/copilot), so Helpdesk's CopilotProcessor applies them
// unchanged. Writes leave CopilotPending / NeedsProcessing / CopilotActions rows that the Helpdesk app finishes with its
// own rules, so "Process Copilot requests" must be on in Helpdesk (Settings > Copilot agent).

import { IAppModule, IAppTarget, Json } from '../types';
import { body, card, confirm, expand, filterParam, fixed, idParam, itemResponse, itemsPath, listGet, listResponse, topParam } from '../openapi';

/** Ticket actions Helpdesk's CopilotProcessor accepts (keep in step with services/copilot/actions.ts; a test checks). */
export const HELPDESK_ACTIONS: Array<{ name: string; payload: string }> = [
  { name: 'setStatus', payload: '{"status":"<name>","comment":"<optional reply>"}' },
  { name: 'resolve', payload: '{"comment":"<optional>"}' },
  { name: 'close', payload: '{"comment":"<optional>"}' },
  { name: 'reopen', payload: '{"comment":"<what is still wrong>"}' },
  { name: 'setPriority', payload: '{"priority":"<name>"}' },
  { name: 'assign', payload: '{"assigneeEmail":"<email>"}' },
  { name: 'assignToMe', payload: '{}' },
  { name: 'unassign', payload: '{}' },
  { name: 'setTeam', payload: '{"team":"<name>"}' },
  { name: 'setCategory', payload: '{"category":"<name>"}' },
  { name: 'addTags', payload: '{"tags":["a","b"]}' },
  { name: 'approve', payload: '{"comment":"<optional>"}' },
  { name: 'reject', payload: '{"comment":"<reason>"}' },
  { name: 'rate', payload: '{"rating":5,"comment":"<optional>"}' },
  { name: 'logTime', payload: '{"minutes":30,"note":"<what>","billable":false}' },
  { name: 'addWatcher', payload: '{"email":"<colleague email>"}' },
  { name: 'merge', payload: '{"target":"<ticket number>"}' }
];

/** Indexed in Helpdesk's schema; Graph rejects filters on anything else. */
export const TICKET_FILTER_FIELDS = ['TicketNumber', 'RequesterEmail', 'AssigneeEmail', 'StatusKind', 'ApprovalStatus', 'SlaResolutionBreached'];

const TICKET = [
  'Title', 'TicketNumber', 'StatusName', 'StatusKind', 'PriorityName', 'CategoryName', 'TeamName', 'RequesterName', 'RequesterEmail',
  'AssigneeName', 'AssigneeEmail', 'ApprovalStatus', 'ResolutionDue', 'SlaResolutionBreached', 'Created', 'Modified'
];
const TICKET_DETAIL = TICKET.concat(['Details', 'TagsText', 'TicketSource', 'FirstResponseDue', 'FirstRespondedAt', 'ResolvedAt', 'SlaResponseBreached', 'CsatRating', 'CsatComment', 'ApprovalComment']);

const ticketPrefix = (t: IAppTarget): string => t.options.ticketPrefix || 'HD-';

export const helpdeskModule: IAppModule = {
  key: 'helpdesk',
  label: 'Helpdesk',
  defaultPrefix: 'HD_',
  namespace: 'helpdesk',
  lists: [
    { key: 'tickets', suffix: 'Tickets' },
    { key: 'comments', suffix: 'Comments' },
    { key: 'kbArticles', suffix: 'KBArticles' },
    { key: 'categories', suffix: 'Categories' },
    { key: 'statuses', suffix: 'Statuses' },
    { key: 'priorities', suffix: 'Priorities' },
    { key: 'agents', suffix: 'Agents' },
    { key: 'copilotActions', suffix: 'CopilotActions' }
  ],

  check(t) {
    return t.options.copilotEnabled === 'true'
      ? []
      : ['In Helpdesk, turn on Settings > Copilot agent > "Process Copilot requests". Without it, tickets and changes made in Copilot are saved but never routed, numbered or applied.'];
  },

  instructions(t) {
    const ex = `${ticketPrefix(t)}00042`;
    return `## Helpdesk (tickets, requests, knowledge)
- Helpdesk agents: call helpdesk.listAgents; the user is an agent if their email is an AgentEmail with IsActive true. Agents may see all tickets and use every action; everyone else only their own requests.
- Ticket numbers look like ${ex}; the item id is the number without prefix and leading zeros (${ex} -> 42).
- searchTickets takes an OData $filter on fields, only on ${TICKET_FILTER_FIELDS.join(', ')}. Examples: fields/RequesterEmail eq 'pat@contoso.com'; open = (fields/StatusKind eq 'New' or fields/StatusKind eq 'Open' or fields/StatusKind eq 'Pending'); fields/SlaResolutionBreached eq 1; fields/ApprovalStatus eq 'Pending'. Never $orderby. Emails lower case. For non-agents ALWAYS add fields/RequesterEmail eq '<their email>'.
- Summaries: getTicket, then listComments with fields/TicketId eq <id>. Never show comments with IsInternal true to non-agents.
- How-to questions: searchKnowledge first (fields/ArticleStatus eq 'Published' and fields/Audience eq 'Everyone' for non-agents); offer an article before raising a ticket.
- Raise: createTicket with Title, Details, CategoryId (from listCategories), RequesterEmail, TicketSource "Copilot", CopilotPending true; PriorityId only if urgent. Reply with the number (${ticketPrefix(t)} + id padded to 5 digits).
- Reply: addComment with Title "Reply on #<id>", TicketId, BodyHtml, CommentKind "Reply", NeedsProcessing true, IsInternal false (true only for agents' internal notes).
- Change a ticket: queueTicketAction with Title = ActionName, TicketId, ActionStatus "Pending", Payload JSON. Actions: ${HELPDESK_ACTIONS.map((a) => `${a.name} ${a.payload}`).join('; ')}. Requesters may resolve, reopen, rate, addWatcher on their own tickets; named approvers approve/reject; agents everything. Applied within about a minute; check with getActionResult.
- Portal: ${t.siteUrl}`;
  },

  paths(t) {
    const g = t.graphSiteId;
    const l = t.listIds;
    const tickets = itemsPath(g, l.tickets);
    const actions = itemsPath(g, l.copilotActions);
    const paths: Json = {
      [tickets]: {
        get: {
          operationId: 'searchTickets',
          summary: 'Search helpdesk tickets with an OData filter on fields.',
          description: "Filter examples: fields/RequesterEmail eq 'a@b.com'; fields/StatusKind eq 'Open'.",
          parameters: [filterParam(`OData filter on fields. Only these fields: ${TICKET_FILTER_FIELDS.join(', ')}.`), fixed('$expand', expand(TICKET)), topParam],
          responses: listResponse('Tickets')
        },
        post: {
          operationId: 'createTicket',
          summary: 'Raise a new helpdesk ticket.',
          description: 'The helpdesk then numbers, routes and notifies.',
          requestBody: body(
            'New ticket',
            {
              Title: { type: 'string', description: 'Short subject' },
              Details: { type: 'string', description: 'Full description' },
              CategoryId: { type: 'integer', description: 'Id from listCategories' },
              PriorityId: { type: 'integer', description: 'Only when urgent; id from listPriorities' },
              RequesterEmail: { type: 'string', description: 'Email of the person the ticket is for' },
              TicketSource: { type: 'string', enum: ['Copilot'], default: 'Copilot' },
              CopilotPending: { type: 'boolean', enum: [true], default: true }
            },
            ['Title', 'Details', 'RequesterEmail', 'TicketSource', 'CopilotPending']
          ),
          responses: itemResponse('Created ticket')
        }
      },
      [`${tickets}/{id}`]: {
        get: {
          operationId: 'getTicket',
          summary: 'One ticket with full details.',
          description: 'Get a ticket by id (number without prefix and leading zeros).',
          parameters: [idParam('Ticket item id'), fixed('$expand', expand(TICKET_DETAIL))],
          responses: itemResponse('Ticket')
        }
      },
      [itemsPath(g, l.comments)]: {
        get: {
          operationId: 'listComments',
          summary: 'Conversation of a ticket.',
          description: 'Replies and internal notes. Never show IsInternal=true rows to non-agents.',
          parameters: [filterParam('fields/TicketId eq <ticket id>'), fixed('$expand', expand(['Title', 'TicketId', 'BodyHtml', 'IsInternal', 'CommentKind', 'Created'])), topParam],
          responses: listResponse('Comments')
        },
        post: {
          operationId: 'addComment',
          summary: 'Reply on a ticket (or an internal note for agents).',
          description: 'The helpdesk then applies first-response, status and notification rules.',
          requestBody: body(
            'Reply',
            {
              Title: { type: 'string' },
              TicketId: { type: 'integer' },
              BodyHtml: { type: 'string', description: 'The reply text' },
              IsInternal: { type: 'boolean' },
              CommentKind: { type: 'string', enum: ['Reply'], default: 'Reply' },
              NeedsProcessing: { type: 'boolean', enum: [true], default: true }
            },
            ['Title', 'TicketId', 'BodyHtml', 'IsInternal', 'CommentKind', 'NeedsProcessing']
          ),
          responses: itemResponse('Reply')
        }
      },
      [actions]: {
        post: {
          operationId: 'queueTicketAction',
          summary: 'Request a change to a ticket.',
          description: 'Applied by the helpdesk within about a minute with its normal rules. Returns an id for getActionResult.',
          requestBody: body(
            'Ticket action',
            {
              Title: { type: 'string' },
              ActionName: { type: 'string', enum: HELPDESK_ACTIONS.map((a) => a.name) },
              TicketId: { type: 'integer' },
              Payload: { type: 'string', description: 'JSON string with the action parameters' },
              ActionStatus: { type: 'string', enum: ['Pending'], default: 'Pending' }
            },
            ['Title', 'ActionName', 'TicketId', 'Payload', 'ActionStatus']
          ),
          responses: itemResponse('Queued action')
        }
      },
      [`${actions}/{id}`]: {
        get: {
          operationId: 'getActionResult',
          summary: 'Result of a queued ticket action.',
          description: 'Pending/Processing = not applied yet; Done = applied; Failed/Rejected = see ResultText.',
          parameters: [idParam('Id returned by queueTicketAction'), fixed('$expand', expand(['Title', 'ActionName', 'TicketId', 'ActionStatus', 'ResultText', 'ProcessedAt']))],
          responses: itemResponse('Action')
        }
      },
      [itemsPath(g, l.kbArticles)]: {
        get: {
          operationId: 'searchKnowledge',
          summary: 'Knowledge base articles.',
          description: "Published articles; for non-agents also fields/Audience eq 'Everyone'.",
          parameters: [filterParam("Default: fields/ArticleStatus eq 'Published' and fields/Audience eq 'Everyone'"), fixed('$expand', expand(['Title', 'Summary', 'BodyHtml', 'TagsText', 'ArticleStatus', 'Audience', 'Featured'])), topParam],
          responses: listResponse('Articles')
        }
      },
      [itemsPath(g, l.categories)]: listGet('listCategories', 'Request categories (item id = CategoryId).', expand(['Title', 'Details', 'RequiresApproval', 'IsActive', 'ShowOnPortal'])),
      [itemsPath(g, l.statuses)]: listGet('listStatuses', 'Ticket statuses and their kind.', expand(['Title', 'StatusKind', 'SortOrder'])),
      [itemsPath(g, l.priorities)]: listGet('listPriorities', 'Priorities and SLA targets.', expand(['Title', 'PriorityLevel', 'ResponseMinutes', 'ResolutionMinutes'])),
      [itemsPath(g, l.agents)]: listGet('listAgents', 'Helpdesk agents and admins.', expand(['Title', 'AgentEmail', 'AgentRole', 'Availability', 'IsActive', 'Skills']))
    };
    return paths;
  },

  functions: [
    {
      name: 'searchTickets',
      description: 'Search helpdesk tickets with an OData filter on fields.',
      capabilities: { response_semantics: { data_path: '$.value', properties: { title: '$.fields.Title', subtitle: '$.fields.TicketNumber' }, static_template: card('${fields.TicketNumber} · ${fields.Title}', [['Status', '${fields.StatusName}'], ['Priority', '${fields.PriorityName}'], ['Assignee', '${fields.AssigneeName}']]) } }
    },
    {
      name: 'getTicket',
      description: 'Get one ticket with full details by id.',
      capabilities: { response_semantics: { data_path: '$', properties: { title: '$.fields.Title', subtitle: '$.fields.TicketNumber' }, static_template: card('${fields.TicketNumber} · ${fields.Title}', [['Status', '${fields.StatusName}'], ['Priority', '${fields.PriorityName}'], ['Assignee', '${fields.AssigneeName}'], ['Requester', '${fields.RequesterName}']]) } }
    },
    { name: 'listComments', description: 'Get the conversation of a ticket.' },
    { name: 'createTicket', description: 'Raise a new helpdesk ticket for the user.', capabilities: confirm('Raise a ticket', 'Raise this request with the service desk?') },
    { name: 'addComment', description: 'Reply on a ticket.', capabilities: confirm('Add a reply', 'Post this reply on the ticket?') },
    { name: 'queueTicketAction', description: 'Change a ticket: status, resolve, close, reopen, priority, assign, team, category, tags, approve, reject, rate, time, watcher, merge.', capabilities: confirm('Change the ticket', 'Submit this change to the helpdesk?') },
    { name: 'getActionResult', description: 'Check whether a queued ticket action was applied.' },
    { name: 'searchKnowledge', description: 'List published helpdesk knowledge base articles.' },
    { name: 'listCategories', description: 'List helpdesk request categories.' },
    { name: 'listStatuses', description: 'List ticket statuses.' },
    { name: 'listPriorities', description: 'List ticket priorities.' },
    { name: 'listAgents', description: 'List helpdesk agents and admins.' }
  ],

  pluginDescription: (org) => ({
    human: `Tickets, replies and knowledge from the ${org} helpdesk.`,
    model: `Use for anything about ${org} helpdesk tickets, IT or service requests, incidents, approvals and how-to knowledge articles.`
  }),

  starters: [
    { title: 'My requests', text: 'What is the status of my helpdesk requests?' },
    { title: 'Get help', text: "I can't connect to the VPN" }
  ]
};
