import { Workbook, Worksheet, DataValidation, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Survey">
    <DataValidation range="A2:A20" type="any" inputTitle="Optional note" inputMessage="Enter any value" />
    <Row><Cell>Response</Cell></Row>
    <Row><Cell /></Row>
  </Worksheet>
</Workbook>;
