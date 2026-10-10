import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="LiteralText">
      <Row>
        <Cell quotePrefix={true}>=SUM(1, 2)</Cell>
      </Row>
    </Worksheet>
  </Workbook>
);
