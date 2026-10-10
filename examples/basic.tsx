import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name={"Sales \"Q\""}>
    <Row><Cell>Item</Cell><Cell>Units</Cell></Row>
    <Row><Cell>Tea</Cell><Cell value={8} /></Row>
    <Row><Cell>Total</Cell><Cell formula={'=IF(A2="Tea",8,0)'} /></Row>
  </Worksheet>
</Workbook>;
