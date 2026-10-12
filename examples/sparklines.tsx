import { Workbook, Worksheet, Row, Cell, SparklineGroup } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Trends">
    <SparklineGroup cells="D2:D3" range="Trends!A2:C3" sparklineType="column" />
    <Row><Cell>Month 1</Cell><Cell>Month 2</Cell><Cell>Month 3</Cell><Cell>Trend</Cell></Row>
    <Row><Cell value={12} /><Cell value={18} /><Cell value={15} /></Row>
    <Row><Cell value={8} /><Cell value={11} /><Cell value={17} /></Row>
  </Worksheet>
</Workbook>;
