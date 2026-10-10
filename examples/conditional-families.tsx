import { Workbook, Worksheet, Row, Cell, ConditionalFormat } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Conditional families">
      <ConditionalFormat range="A2:A6" type="average" rule="twoStdDevBelow" fillColor="#FFC7CE" strikethrough={true} underline="double" script="superscript" numberFormat="0.00" borderTop="thin" borderTopColor="#123456" />
      <ConditionalFormat range="B2:B6" type="date" rule="thisMonth" fontColor="#9C0006" />
      <ConditionalFormat range="C2:C6" type="blank" fillColor="#FFF2CC" />
      <ConditionalFormat range="D2:D6" type="error" rule="nonError" bold={true} />
      <ConditionalFormat range="E2:E6" type="duplicate" rule="unique" italic={true} />
      <Row><Cell>Average</Cell><Cell>Date</Cell><Cell>Blank</Cell><Cell>Error</Cell><Cell>Unique</Cell></Row>
      <Row><Cell value={12} /><Cell value="2026-10-11" /><Cell>present</Cell><Cell value={1} /><Cell>alpha</Cell></Row>
      <Row><Cell value={18} /><Cell value="2026-10-12" /><Cell /><Cell value={2} /><Cell>beta</Cell></Row>
      <Row><Cell value={4} /><Cell value="2026-10-13" /><Cell>present</Cell><Cell value={3} /><Cell>alpha</Cell></Row>
      <Row><Cell value={9} /><Cell value="2026-10-14" /><Cell /><Cell value={4} /><Cell>delta</Cell></Row>
      <Row><Cell value={21} /><Cell value="2026-10-15" /><Cell>present</Cell><Cell value={5} /><Cell>epsilon</Cell></Row>
    </Worksheet>
  </Workbook>
);
