import { Workbook, Worksheet, ConditionalFormat, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Scores">
      <ConditionalFormat
        range="B2:B20"
        type="colorScale2"
        minColor="#FFF2CC"
        maxColor="#70AD47"
        minType="number"
        minValue={0}
        maxType="number"
        maxValue={100}
      />
      <ConditionalFormat
        range="C2:C20"
        type="colorScale3"
        minColor="#F8696B"
        midColor="#FFEB84"
        maxColor="#63BE7B"
        midType="percentile"
        midValue={50}
      />
      <Row><Cell>Student</Cell><Cell>Score</Cell><Cell>Growth</Cell></Row>
      <Row><Cell>Ada</Cell><Cell value={92} /><Cell value={14} /></Row>
    </Worksheet>
  </Workbook>
);
