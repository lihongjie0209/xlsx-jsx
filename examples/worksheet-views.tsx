import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Page layout" view="pageLayout" zoom={125}>
    <Row><Cell>Page layout view</Cell></Row>
  </Worksheet>
  <Worksheet name="Page breaks" view="pageBreakPreview" zoom={90}>
    <Row><Cell>Page break preview</Cell></Row>
  </Worksheet>
</Workbook>;
