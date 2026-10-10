import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Shrink to fit">
    <Row>
      <Cell shrinkToFit={true} horizontalAlign="right">
        Compact this content to fit the cell width.
      </Cell>
      <Cell>Normal width</Cell>
    </Row>
  </Worksheet>
</Workbook>;
