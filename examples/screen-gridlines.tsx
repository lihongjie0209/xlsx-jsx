import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Presentation" showGridlines={false}>
    <Row><Cell>On-screen gridlines are hidden.</Cell></Row>
  </Worksheet>
</Workbook>;
