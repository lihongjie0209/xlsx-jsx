import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Links">
      <Row>
        <Cell hyperlink="https://example.com/reports?q=weekly" hyperlinkTooltip="View the latest report">Open weekly report</Cell>
        <Cell hyperlink="mailto:reports@example.com">Email the report team</Cell>
        <Cell hyperlink="internal:Reports!A1" hyperlinkTooltip="Jump to the report sheet">Open the report sheet</Cell>
      </Row>
    </Worksheet>
    <Worksheet name="Reports">
      <Row><Cell>Weekly report</Cell></Row>
    </Worksheet>
  </Workbook>
);
