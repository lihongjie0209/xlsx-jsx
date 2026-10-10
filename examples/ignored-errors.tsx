import { Workbook, Worksheet, IgnoreError, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <IgnoreError range="A2:A20" errorType="numberStoredAsText" />
    <IgnoreError range="C2:C20" errorType="evalError" />
    <Row><Cell>Account</Cell><Cell>Amount</Cell><Cell>Formula</Cell></Row>
  </Worksheet>
</Workbook>;
