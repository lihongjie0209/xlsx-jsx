import { Workbook, Worksheet, Row, Cell, Sparkline } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Trends">
    <Sparkline cell="D2" range="Trends!A2:C2" style={2} sparklineColor="#4472C4" />
    <Row><Cell value={12} /><Cell value={18} /><Cell value={15} /><Cell>Trend</Cell></Row>
  </Worksheet>
</Workbook>;
