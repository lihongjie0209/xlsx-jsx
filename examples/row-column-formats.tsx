import { Workbook, Worksheet, Column, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Defaults">
    <Column min={2} max={2} numberFormat="$#,##0.00" horizontalAlign="right" />
    <Row bold={true} fillColor="#DDEEFF">
      <Cell>Item</Cell><Cell>Amount</Cell>
    </Row>
    <Row>
      <Cell>Revenue</Cell><Cell value={125000} />
    </Row>
  </Worksheet>
</Workbook>;
