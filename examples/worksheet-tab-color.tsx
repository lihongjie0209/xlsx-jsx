import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Quarterly results" tabColor={{ theme: 4, shade: 2 }} codeName="QuarterlyResults">
    <Row><Cell>Blue identifies the reporting worksheet.</Cell></Row>
  </Worksheet>
</Workbook>;
