import { Workbook, Worksheet, Row, Cell, Column } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Defaults" defaultColumnWidthPixels={80}>
    <Column min={2} max={2} width={18.5} />
    <Row>
      <Cell>Default width is 80 pixels</Cell>
      <Cell>Explicit column width overrides it</Cell>
    </Row>
  </Worksheet>
</Workbook>;
