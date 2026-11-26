// Brillienta Asset Management in the suite agent.
//
// Reads: my equipment, who has a tag, equipment requests, asset types.
// Writes (when the site supports them):
//   - requests (new equipment, loans, report a fault, return, cancel) are queued in <prefix>CopilotActions and applied
//     by Asset Management's own CopilotProcessor with the portal's rules, as the person who asked (Created By). Needs
//     Asset Management with that list (schema 3) and Settings > General > Features > "Requests from Microsoft 365
//     Copilot" on;
//   - confirming receipt writes <prefix>Acknowledgements, which Asset Management already processes the same way.
// Otherwise the module is read-only and the setup page says why.

import { IAppModule, IAppTarget, IFunctionSpec } from '../types';
import { body, card, confirm, expand, filterParam, fixed, idParam, itemResponse, itemsPath, listGet, listResponse, topParam } from '../openapi';

/** Indexed in Asset Management's schema; Graph rejects filters on anything else. */
export const ASSET_FILTER_FIELDS = ['AssetTag', 'SerialNumber', 'AssignedToEmail', 'StatusKind', 'AssetTypeId', 'LocationId'];
export const REQUEST_FILTER_FIELDS = ['RequestNumber', 'RequesterEmail', 'RequestStatus', 'ApprovalStatus', 'AssetId'];

/** Actions Asset Management's CopilotProcessor accepts (keep in step with services/copilotActions.ts; a test checks). */
export const ASSET_ACTIONS: Array<{ name: string; payload: string }> = [
  { name: 'requestAsset', payload: '{"assetType":"Laptop","details":"<what and why>","kind":"New|Loan","loanUntil":"yyyy-mm-dd (loans)","neededBy":"yyyy-mm-dd","urgency":"Low|Normal|High|Urgent"}' },
  { name: 'reportIssue', payload: '{"assetTag":"<their asset>","details":"<what is wrong>"}' },
  { name: 'returnAsset', payload: '{"assetTag":"<their asset>","details":"<optional>"}' },
  { name: 'cancelRequest', payload: '{"requestNumber":"<their open request>","reason":"<optional>"}' }
];

const ASSET = [
  'Title', 'AssetTag', 'AssetTypeName', 'StatusName', 'StatusKind', 'Manufacturer', 'ModelName', 'SerialNumber', 'AssignedToName', 'AssignedToEmail',
  'AssignedAt', 'AckState', 'LocationName', 'Department', 'Condition', 'DeviceName', 'OperatingSystem', 'ComplianceState', 'LastAuditAt'
];
const REQUEST = [
  'Title', 'RequestNumber', 'RequestKind', 'AssetTypeName', 'AssetTag', 'RequesterName', 'RequesterEmail', 'RequestStatus', 'ApprovalStatus',
  'ApproverName', 'Urgency', 'LocationName', 'Created', 'FulfilledAt', 'ResolutionNote'
];

/** Requests can be queued: the list exists and Asset Management applies them. */
export const canRequest = (t: IAppTarget): boolean => !!t.listIds.copilotActions && t.options.copilotEnabled === 'true';
/** Receipt can be confirmed: the list exists and the feature is on. */
export const canConfirm = (t: IAppTarget): boolean => !!t.listIds.acknowledgements && t.options.acknowledgementsEnabled !== 'false';

const READ_FUNCTIONS: IFunctionSpec[] = [
  {
    name: 'searchAssets',
    description: 'Search equipment (laptops, phones, monitors...) by who has it, tag, serial number or status.',
    capabilities: { response_semantics: { data_path: '$.value', properties: { title: '$.fields.Title', subtitle: '$.fields.AssetTag' }, static_template: card('${fields.AssetTag} · ${fields.Title}', [['Type', '${fields.AssetTypeName}'], ['Status', '${fields.StatusName}'], ['Assigned to', '${fields.AssignedToName}']]) } }
  },
  { name: 'getAsset', description: 'Get one asset with details by id.' },
  { name: 'searchAssetRequests', description: 'Search equipment requests by requester or status.' },
  { name: 'listAssetTypes', description: 'List asset types and which ones can be requested.' }
];
const REQUEST_FUNCTIONS: IFunctionSpec[] = [
  { name: 'queueAssetAction', description: 'Request equipment, borrow it, report a fault, return it, or cancel an open request, for the signed-in person.', capabilities: confirm('Send to the asset team', 'Submit this to the asset team?') },
  { name: 'getAssetActionResult', description: 'Check whether a queued equipment request was applied, and its request number.' }
];
const CONFIRM_FUNCTIONS: IFunctionSpec[] = [
  { name: 'confirmReceipt', description: 'Confirm (or dispute) receipt of equipment assigned to the signed-in person.', capabilities: confirm('Confirm receipt', 'Record this answer about the equipment you received?') }
];

export const assetsModule: IAppModule = {
  key: 'assets',
  label: 'Asset Management',
  defaultPrefix: 'BrillAM_',
  namespace: 'assets',
  lists: [
    { key: 'assets', suffix: 'Assets' },
    { key: 'requests', suffix: 'Requests' },
    { key: 'assetTypes', suffix: 'AssetTypes' },
    { key: 'copilotActions', suffix: 'CopilotActions', optional: true },
    { key: 'acknowledgements', suffix: 'Acknowledgements', optional: true }
  ],

  notes(t) {
    const out: string[] = [];
    if (!t.listIds.copilotActions) out.push('Requests are read-only: update Asset Management to the version with Copilot requests to raise, return and report from Copilot.');
    else if (t.options.copilotEnabled !== 'true') out.push('Requests are read-only: in Asset Management, turn on Settings > General > Features > "Requests from Microsoft 365 Copilot".');
    return out;
  },

  instructions(t) {
    const lines = [
      '## Asset Management (equipment)',
      `- searchAssets takes an OData $filter on fields, only on ${ASSET_FILTER_FIELDS.join(', ')}. My equipment: fields/AssignedToEmail eq '<my email, lower case>'. Who has a tag: fields/AssetTag eq 'LT-0142'. StatusKind is Available, Assigned, Maintenance, Retired, Disposed or Lost.`,
      "- searchAssetRequests: fields/RequesterEmail eq '<email>'; RequestStatus is Submitted, PendingApproval, Approved, Rejected, InProgress, Fulfilled, Cancelled or Closed.",
      '- Show equipment as a table: tag, type, model, status, assigned to. AckState Pending means the person has not confirmed receipt yet.'
    ];
    if (canRequest(t)) {
      lines.push(
        `- To request, borrow, report a fault, return or cancel: queueAssetAction with Title = ActionName, ActionStatus "Pending" and Payload JSON. Actions: ${ASSET_ACTIONS.map((a) => `${a.name} ${a.payload}`).join('; ')}. assetType is a name from listAssetTypes (Requestable true). Only for the signed-in person and their own assets and requests. The asset team's app applies it within about a minute with its approvals; then call getAssetActionResult with the returned id and tell them the request number from ResultText.`
      );
    } else {
      lines.push(`- You cannot raise, return or report equipment here: send the user to the Asset Management portal: ${t.siteUrl}.`);
    }
    if (canConfirm(t)) {
      lines.push("- When someone says they received (or did not receive) equipment assigned to them with AckState Pending: confirmReceipt with Title = the asset tag, AssetId = the asset's item id, AckResponse Accepted or Disputed, AckComment, AckStatus \"Pending\".");
    }
    return lines.join('\n');
  },

  paths(t) {
    const g = t.graphSiteId;
    const l = t.listIds;
    const assets = itemsPath(g, l.assets);
    const out: Record<string, unknown> = {
      [assets]: {
        get: {
          operationId: 'searchAssets',
          summary: 'Search equipment with an OData filter on fields.',
          description: "Filter examples: fields/AssignedToEmail eq 'a@b.com'; fields/AssetTag eq 'LT-0142'.",
          parameters: [filterParam(`OData filter on fields. Only these fields: ${ASSET_FILTER_FIELDS.join(', ')}.`), fixed('$expand', expand(ASSET)), topParam],
          responses: listResponse('Assets')
        }
      },
      [`${assets}/{id}`]: {
        get: {
          operationId: 'getAsset',
          summary: 'One asset with details.',
          description: 'Get an asset by its list item id.',
          parameters: [idParam('Asset item id'), fixed('$expand', expand(ASSET.concat(['PurchaseCost', 'VendorName', 'TagsText'])))],
          responses: itemResponse('Asset')
        }
      },
      [itemsPath(g, l.requests)]: {
        get: {
          operationId: 'searchAssetRequests',
          summary: 'Search equipment requests with an OData filter on fields.',
          description: "Filter example: fields/RequesterEmail eq 'a@b.com'.",
          parameters: [filterParam(`OData filter on fields. Only these fields: ${REQUEST_FILTER_FIELDS.join(', ')}.`), fixed('$expand', expand(REQUEST)), topParam],
          responses: listResponse('Requests')
        }
      },
      [itemsPath(g, l.assetTypes)]: listGet('listAssetTypes', 'Asset types and whether employees can request them.', expand(['Title', 'Details', 'Requestable', 'RequiresApproval', 'AllowLoans', 'IsActive']))
    };
    if (canRequest(t)) {
      const queue = itemsPath(g, l.copilotActions);
      out[queue] = {
        post: {
          operationId: 'queueAssetAction',
          summary: 'Send an equipment request to the asset team.',
          description: 'Asset Management applies it within about a minute with its normal rules and approvals. Returns an id for getAssetActionResult.',
          requestBody: body(
            'Equipment action',
            {
              Title: { type: 'string', description: 'Same as ActionName' },
              ActionName: { type: 'string', enum: ASSET_ACTIONS.map((a) => a.name) },
              Payload: { type: 'string', description: 'JSON string with the action details' },
              ActionStatus: { type: 'string', enum: ['Pending'], default: 'Pending' }
            },
            ['Title', 'ActionName', 'Payload', 'ActionStatus']
          ),
          responses: itemResponse('Queued action')
        }
      };
      out[`${queue}/{id}`] = {
        get: {
          operationId: 'getAssetActionResult',
          summary: 'Result of a queued equipment request.',
          description: 'Pending/Processing = not applied yet; Done = applied (ResultText has the request number); Failed/Rejected = see ResultText.',
          parameters: [idParam('Id returned by queueAssetAction'), fixed('$expand', expand(['Title', 'ActionName', 'ActionStatus', 'ResultText', 'RequestId', 'ProcessedAt']))],
          responses: itemResponse('Action')
        }
      };
    }
    if (canConfirm(t)) {
      out[itemsPath(g, l.acknowledgements)] = {
        post: {
          operationId: 'confirmReceipt',
          summary: 'Confirm or dispute receipt of assigned equipment.',
          description: 'Recorded on the asset by the asset team\'s app; only the person the asset is assigned to can confirm it.',
          requestBody: body(
            'Receipt confirmation',
            {
              Title: { type: 'string', description: 'The asset tag' },
              AssetId: { type: 'integer', description: 'The asset item id' },
              AckResponse: { type: 'string', enum: ['Accepted', 'Disputed'] },
              AckComment: { type: 'string', description: 'Optional; required in spirit when disputing' },
              AckStatus: { type: 'string', enum: ['Pending'], default: 'Pending' }
            },
            ['Title', 'AssetId', 'AckResponse', 'AckStatus']
          ),
          responses: itemResponse('Confirmation')
        }
      };
    }
    return out;
  },

  functions: READ_FUNCTIONS.concat(REQUEST_FUNCTIONS, CONFIRM_FUNCTIONS),

  functionsFor(t) {
    return READ_FUNCTIONS.concat(canRequest(t) ? REQUEST_FUNCTIONS : [], canConfirm(t) ? CONFIRM_FUNCTIONS : []);
  },

  pluginDescription: (org) => ({
    human: `Equipment and equipment requests from ${org} Asset Management.`,
    model: `Use for ${org} IT equipment and assets: laptops, phones, monitors, who has which device, asset tags, serial numbers, requesting, borrowing, returning or reporting equipment, and confirming receipt.`
  }),

  starters: [{ title: 'My equipment', text: 'What equipment is assigned to me?' }]
};
