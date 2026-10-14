export type ProgressInput = {
  id?: number;
  memberName: string;
  month: string;
  tasks: number;
  observations: number;
};

export function validateProgress(value: unknown): ProgressInput | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const memberName =
    typeof row.memberName === "string"
      ? row.memberName.normalize("NFKC").replace(/\s+/g, " ").trim()
      : "";
  const month = typeof row.month === "string" ? row.month.trim() : "";
  const tasks = Number(row.tasks);
  const observations = Number(row.observations);
  if (
    !memberName || memberName.length > 100 ||
    !/^\d{4}-(0[1-9]|1[0-2])$/.test(month) ||
    !Number.isInteger(tasks) || tasks < 0 || tasks > 999999 ||
    !Number.isInteger(observations) || observations < 0 || observations > 999999
  ) return null;
  const id = row.id === undefined ? undefined : Number(row.id);
  if (id !== undefined && (!Number.isInteger(id) || id <= 0)) return null;
  return { id, memberName, month, tasks, observations };
}

export function nameKey(name: string) {
  return name.toLocaleLowerCase();
}
