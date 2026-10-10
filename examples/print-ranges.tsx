import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet
      name="Monthly Report"
      printArea="A1:D20"
      repeatRows="1:2"
      repeatColumns="A:A"
    >
      <Row><Cell>Month</Cell><Cell>Revenue</Cell><Cell>Costs</Cell><Cell>Net</Cell></Row>
      <Row><Cell>Units</Cell><Cell>USD</Cell><Cell>USD</Cell><Cell>USD</Cell></Row>
      <Row><Cell>January</Cell><Cell value={1200} /><Cell value={800} /><Cell value={400} /></Row>
    </Worksheet>
  </Workbook>
);
