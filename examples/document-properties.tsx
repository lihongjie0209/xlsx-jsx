import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook
  title="Quarterly report"
  subject="Revenue"
  author="Finance team"
  keywords="quarterly,finance"
  description="Reviewed & approved"
  category="Finance"
  status="Draft"
  manager="Ada Lovelace"
  company="Analytical Engines"
  hyperlinkBase="https://example.test/reports/"
>
  <Worksheet name="Summary">
    <Row><Cell>Revenue</Cell><Cell value={125000} /></Row>
  </Worksheet>
</Workbook>;
