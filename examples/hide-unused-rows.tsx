import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Compact view" hideUnusedRows={true}>
    <Row><Cell>Rows without content are hidden.</Cell></Row>
  </Worksheet>
</Workbook>;
