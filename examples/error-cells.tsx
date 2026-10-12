import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Errors">
      <Row>
        <Cell error="#NULL!" />
        <Cell error="#DIV/0!" />
        <Cell error="#VALUE!" />
        <Cell error="#REF!" />
        <Cell error="#NAME?" />
        <Cell error="#NUM!" />
        <Cell error="#N/A" />
      </Row>
    </Worksheet>
  </Workbook>
);
