import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Printed report" printGridlines={true}>
    <Row><Cell>Gridlines appear on printed pages.</Cell></Row>
  </Worksheet>
</Workbook>;
