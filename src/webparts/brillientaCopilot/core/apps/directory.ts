// Brillienta Employee Directory in the suite agent. Read-only: departments, offices, managers and direct reports.
// Only the fields a directory card shows are requested (never mobile phone, hire date or employee number), and
// people hidden from the directory or inactive are never shown.

import { IAppModule } from '../types';
import { card, expand, filterParam, fixed, itemsPath, listGet, listResponse, topParam } from '../openapi';

/** Indexed in Employee Directory's schema; Graph rejects filters on anything else. */
export const PEOPLE_FILTER_FIELDS = ['AadId', 'Mail', 'Department', 'OfficeLocation', 'Country', 'ManagerAadId', 'EmpStatus'];

/** Card fields only. Deliberately excludes MobilePhone, HireDate, EmployeeNumber, GraphExtras and Overrides. */
export const PEOPLE_FIELDS = [
  'Title', 'GivenName', 'Surname', 'Mail', 'JobTitle', 'Department', 'OfficeLocation', 'City', 'Country', 'BusinessPhone',
  'ManagerAadId', 'AadId', 'EmpStatus', 'UserKind', 'HiddenFromDirectory'
];

export const directoryModule: IAppModule = {
  key: 'directory',
  label: 'Employee Directory',
  defaultPrefix: 'BrillED_',
  namespace: 'directory',
  lists: [
    { key: 'employees', suffix: 'Employees' },
    { key: 'departments', suffix: 'Departments' }
  ],

  instructions(t) {
    return `## Employee Directory (people, read-only)
- To find a person by name use your built-in people search. Use searchPeople for departments, offices, countries, managers and direct reports.
- searchPeople takes an OData $filter on fields, only on ${PEOPLE_FILTER_FIELDS.join(', ')}. Always add fields/EmpStatus eq 'Active'. Me: fields/AadId eq '<id from getCurrentUser>'. My manager: read my ManagerAadId, then fields/AadId eq '<it>'. My team (direct reports): fields/ManagerAadId eq '<my id>'. A department: fields/Department eq 'Finance'.
- Never show anyone with HiddenFromDirectory true. Show people as a table: name, job title, department, office, email.
- Profiles, photos and kudos are in the directory: ${t.siteUrl}`;
  },

  paths(t) {
    const g = t.graphSiteId;
    const l = t.listIds;
    return {
      [itemsPath(g, l.employees)]: {
        get: {
          operationId: 'searchPeople',
          summary: 'Search the employee directory with an OData filter on fields.',
          description: "Filter examples: fields/Department eq 'Finance' and fields/EmpStatus eq 'Active'; fields/ManagerAadId eq '<id>'.",
          parameters: [filterParam(`OData filter on fields. Only these fields: ${PEOPLE_FILTER_FIELDS.join(', ')}.`), fixed('$expand', expand(PEOPLE_FIELDS)), topParam],
          responses: listResponse('People')
        }
      },
      [itemsPath(g, l.departments)]: listGet('listDepartments', 'Departments with their description and contact.', expand(['Title', 'Details', 'ParentDepartment', 'DeptEmail']))
    };
  },

  functions: [
    {
      name: 'searchPeople',
      description: 'Find colleagues by department, office, country, manager or direct reports.',
      capabilities: { response_semantics: { data_path: '$.value', properties: { title: '$.fields.Title', subtitle: '$.fields.JobTitle' }, static_template: card('${fields.Title}', [['Job title', '${fields.JobTitle}'], ['Department', '${fields.Department}'], ['Office', '${fields.OfficeLocation}'], ['Email', '${fields.Mail}']]) } }
    },
    { name: 'listDepartments', description: 'List departments.' }
  ],

  pluginDescription: (org) => ({
    human: `People, teams and departments from the ${org} employee directory.`,
    model: `Use for questions about ${org} colleagues by department, office, country, manager or direct reports, and about departments.`
  }),

  starters: [{ title: 'My team', text: 'Who is in my team and who is my manager?' }]
};
