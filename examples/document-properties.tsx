import { Workbook, Worksheet, Row, Cell, CustomProperty } from "xlsx-jsx";

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
  codeName="QuarterlyReport"
>
  <Worksheet name="Summary">
    <Row><Cell>Revenue</Cell><Cell value={125000} /></Row>
  </Worksheet>
  <CustomProperty name="Department" type="string" value="Finance & Risk" />
  <CustomProperty name="Document number" type="integer" value={5321} />
  <CustomProperty name="Confidence" type="real" value={0.875} />
  <CustomProperty name="Approved" type="boolean" value={true} />
  <CustomProperty name="Reviewed at" type="dateTime" value="2026-10-12T09:30:00Z" />
</Workbook>;
