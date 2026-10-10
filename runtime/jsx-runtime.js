function node(type, props = {}) {
  const { children, ...rest } = props;
  return { __xlsxJsx: true, type, props: rest, children: children === undefined ? [] : children };
}
function host(type) { return props => node(type, props ?? {}); }
export const Workbook = host("Workbook");
export const DefinedName = host("DefinedName");
export const Worksheet = host("Worksheet");
export const Column = host("Column");
export const DataValidation = host("DataValidation");
export const ConditionalFormat = host("ConditionalFormat");
export const IgnoreError = host("IgnoreError");
export const Table = host("Table");
export const TableColumn = host("TableColumn");
export const Note = host("Note");
export const Chart = host("Chart");
export const Series = host("Series");
export const Image = host("Image");
export const Row = host("Row");
export const Cell = host("Cell");
export const TextRun = host("TextRun");
export const Fragment = globalThis.Symbol.for("xlsx-jsx.fragment");
export function jsx(type, props) {
  if (type === Fragment) return props?.children ?? [];
  if (typeof type !== "function") throw new TypeError(`Unknown JSX element: ${String(type)}`);
  return type(props ?? {});
}
export const jsxs = jsx;
export const jsxDEV = jsx;
async function normalize(value, parentType) {
  value = await value;
  if (value === null || value === undefined) return [];
  if (typeof value === "boolean") return parentType === "Cell" ? [value] : [];
  if (Array.isArray(value)) {
    const output = [];
    for (const child of value) output.push(...await normalize(child, parentType));
    return output;
  }
  if (typeof value === "string" || typeof value === "number") {
    if (typeof value === "number" && !Number.isFinite(value)) throw new TypeError("Cell numbers must be finite");
    if ((parentType === "Workbook" || parentType === "Worksheet" || parentType === "Row") && typeof value === "string" && value.trim() === "") return [];
    return [value];
  }
  if (!value || value.__xlsxJsx !== true) throw new TypeError("Component returned a value that is not an xlsx-jsx node");
  const children = await normalize(value.children, value.type);
  return [{ type: value.type, props: value.props ?? {}, children }];
}
export async function finalize(value, data) {
  if (typeof value === "function") value = value(data);
  const roots = await normalize(value, undefined);
  if (roots.length !== 1) throw new TypeError("The default export must produce exactly one Workbook");
  return { version: 1, workbook: roots[0] };
}
