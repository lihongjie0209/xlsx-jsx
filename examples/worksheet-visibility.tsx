import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Summary">
      <Row><Cell>Visible worksheet</Cell></Row>
    </Worksheet>
    <Worksheet name="Intermediate" visibility="hidden">
      <Row><Cell>Hidden calculations</Cell></Row>
    </Worksheet>
    <Worksheet name="Internal" visibility="veryHidden">
      <Row><Cell>Internal data</Cell></Row>
    </Worksheet>
  </Workbook>
);
