import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Overview">
    <Row><Cell>Summary</Cell></Row>
  </Worksheet>
  <Worksheet name="Details" active={true} firstTab={true}>
    <Row><Cell>This worksheet opens as the active, leftmost tab.</Cell></Row>
  </Worksheet>
</Workbook>;
