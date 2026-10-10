import { Workbook, Worksheet, ConditionalFormat, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Scores">
      <ConditionalFormat
        range="B2:B20"
        type="cellIs"
        operator="between"
        formula1="60"
        formula2="100"
        bold={true}
        fontColor="#006100"
        fillColor="#C6EFCE"
        stopIfTrue={true}
      />
      <Row><Cell>Student</Cell><Cell>Score</Cell></Row>
      <Row><Cell>Ada</Cell><Cell value={92} /></Row>
    </Worksheet>
  </Workbook>
);
