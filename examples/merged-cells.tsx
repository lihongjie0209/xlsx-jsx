import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Merged cells">
      <Row>
        <Cell rowSpan={2} columnSpan={2} bold={true} horizontalAlign="center">
          Quarterly summary
        </Cell>
        <Cell>Q1</Cell>
        <Cell>Q2</Cell>
      </Row>
      <Row>
        <Cell value={125} numberFormat="#,##0" />
        <Cell value={240} numberFormat="#,##0" />
      </Row>
    </Worksheet>
  </Workbook>
);
