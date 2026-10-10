import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Arabic report" rightToLeft={true}>
    <Row><Cell>البيانات المالية</Cell><Cell>التاريخ</Cell></Row>
    <Row><Cell value={1250} /><Cell>2026-10-10</Cell></Row>
  </Worksheet>
</Workbook>;
