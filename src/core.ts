// Portable core logic with no host-specific types

export interface Entry {
  id?: string;
  date: string; // ISO date string
  notes?: string;
  hours?: number;
}

export type FunctionStringOptions =
  | "Add Entry"
  | "Shift Entry"
  | "Pop Entry"
  | "Clock In"
  | "Clock Out";

export function roundToNearest30(minutes: number): number {
  const remainder = minutes % 30;
  if (remainder === 0) return minutes;
  return minutes + (remainder >= 15 ? 30 - remainder : -remainder);
}

export function addEntry(entries: Entry[], entry: Entry): Entry[] {
  return [...entries, entry];
}

export function popEntry(entries: Entry[]): Entry[] {
  return entries.slice(0, -1);
}

export function shiftEntry(entries: Entry[]): Entry[] {
  if (entries.length === 0) return entries;
  const first = entries[0];
  if (!first) return entries;
  return [...entries.slice(1), first];
}

export function clockIn(entries: Entry[], timestamp: string): Entry[] {
  // Example: add an entry marking clock in time
  const entry: Entry = { date: timestamp };
  return addEntry(entries, entry);
}

export function clockOut(entries: Entry[], timestamp: string): Entry[] {
  // Example placeholder — real logic depends on your data model
  if (entries.length === 0) return entries;
  const lastIndex = entries.length - 1;
  const last = entries[lastIndex];
  if (!last) return entries;
  const updated: Entry = {
    ...last,
    notes: last.notes ? `${last.notes} (clock out ${timestamp})` : `clock out ${timestamp}`,
  };
  return [...entries.slice(0, -1), updated];
}
