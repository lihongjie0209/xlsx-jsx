import { Workbook, Worksheet, Table, TableColumn, ConditionalFormat, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Fills">
    <Table range="A1:A2" name="PatternFills">
      <TableColumn format={{"fillPattern":"lightTrellis","fillForegroundColor":"#336699","fillBackgroundColor":"#DDEEFF"}} />
    </Table>
    <ConditionalFormat range="A2:A3" type="cellIs" operator="greaterThan" formula1="0" fillPattern="darkGrid" fillForegroundColor="#FF0000" fillBackgroundColor="#FFFF00" />
    <Row>
      <Cell fillColor="#DDEEFF">Solid RGB fill</Cell>
      <Cell value={42} fillColor="#E2F0D9" bold={true} />
      <Cell formula="=1+1" fillColor="#FFF2CC" numberFormat="0.00" />
      <Cell fillPattern="darkDown" fillForegroundColor="#FF0000" fillBackgroundColor="#FFFF00">Pattern fill</Cell>
    </Row>
    <Row><Cell value={42} /></Row>
  </Worksheet>
</Workbook>;
