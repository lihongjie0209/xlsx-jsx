import { Workbook, Worksheet, Row, Cell, ConditionalFormat } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Data bars">
      <ConditionalFormat
        range="B2:B6"
        type="dataBar"
        minType="number"
        minValue={0}
        maxType="number"
        maxValue={100}
        fillColor="#4472C4"
        negativeFillColor="#C00000"
        solidFill={true}
        direction="leftToRight"
        axisPosition="midpoint"
      />
      <Row><Cell>Metric</Cell><Cell>Value</Cell></Row>
      <Row><Cell>Alpha</Cell><Cell value={18} /></Row>
      <Row><Cell>Beta</Cell><Cell value={42} /></Row>
      <Row><Cell>Gamma</Cell><Cell value={67} /></Row>
      <Row><Cell>Delta</Cell><Cell value={93} /></Row>
      <Row><Cell>Variance</Cell><Cell value={-12} /></Row>
    </Worksheet>
  </Workbook>
);
