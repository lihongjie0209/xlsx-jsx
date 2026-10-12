import { Workbook, Worksheet, Row, Cell, SparklineGroup } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Trends">
    <SparklineGroup cells="D2:D3" range="Trends!A2:C3" dateRange="Trends!A1:C1" sparklineType="column" style={20} lineWeight={1.5} min={0} max={20} />
    <Row><Cell date="2026-01-01" /><Cell date="2026-02-01" /><Cell date="2026-03-01" /><Cell>Trend</Cell></Row>
    <Row><Cell value={12} /><Cell value={18} /><Cell value={15} /></Row>
    <Row><Cell value={8} /><Cell value={11} /><Cell value={17} /></Row>
  </Worksheet>
</Workbook>;
