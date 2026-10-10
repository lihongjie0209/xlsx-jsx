import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Rows" defaultRowHeightPixels={24}>
    <Row><Cell>Row height is 24 pixels by default.</Cell></Row>
  </Worksheet>
</Workbook>;
