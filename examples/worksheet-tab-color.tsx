import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Quarterly results" tabColor="#4472C4" codeName="QuarterlyResults">
    <Row><Cell>Blue identifies the reporting worksheet.</Cell></Row>
  </Worksheet>
</Workbook>;
