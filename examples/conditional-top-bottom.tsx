import { Workbook, Worksheet, Row, Cell, ConditionalFormat } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Top and bottom">
      <ConditionalFormat range="B2:B6" type="top10" rule="top" rank={2} bold={true} fillColor="#E2F0D9" />
      <ConditionalFormat range="B2:B6" type="top10" rule="bottomPercent" rank={25} fontColor="#9C0006" fillColor="#FFC7CE" stopIfTrue={true} />
      <Row><Cell>Item</Cell><Cell>Score</Cell></Row>
      <Row><Cell>Alpha</Cell><Cell value={92} /></Row>
      <Row><Cell>Beta</Cell><Cell value={81} /></Row>
      <Row><Cell>Gamma</Cell><Cell value={67} /></Row>
      <Row><Cell>Delta</Cell><Cell value={48} /></Row>
      <Row><Cell>Epsilon</Cell><Cell value={35} /></Row>
    </Worksheet>
  </Workbook>
);
