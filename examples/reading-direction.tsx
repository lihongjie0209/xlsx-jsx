import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Direction">
      <Row>
        <Cell readingDirection="leftToRight">English</Cell>
        <Cell readingDirection="rightToLeft">العربية</Cell>
      </Row>
    </Worksheet>
  </Workbook>
);
