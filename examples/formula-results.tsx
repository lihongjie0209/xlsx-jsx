import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Formula results">
      <Row>
        <Cell formula="=1+1" formulaResult={2} />
        <Cell formula={'="ready"'} formulaResult="ready" />
        <Cell formula="=TRUE()" formulaResult={true} />
        <Cell formula="=NA()" formulaResult="#N/A" />
        <Cell formula="=A1" formulaResult={null} />
      </Row>
    </Worksheet>
  </Workbook>
);
