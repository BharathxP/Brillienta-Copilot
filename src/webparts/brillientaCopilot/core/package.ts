// Zips the generated agent files and the two icons into the app package an admin uploads to Microsoft 365.

import { AgentFiles } from './types';
import { createZip, utf8 } from './zip';

export function buildPackageZip(files: AgentFiles, icons: { color: Uint8Array; outline: Uint8Array }, date?: Date): Uint8Array {
  const entries = Object.keys(files)
    .sort()
    .map((name) => ({ name, data: utf8(files[name]) }));
  entries.push({ name: 'color.png', data: icons.color }, { name: 'outline.png', data: icons.outline });
  return createZip(entries, date);
}
