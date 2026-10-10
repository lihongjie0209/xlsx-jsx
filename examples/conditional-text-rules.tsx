import { Workbook, Worksheet, Row, Cell, ConditionalFormat } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Text rules">
      <ConditionalFormat
        range="A2:A5"
        type="text"
        operator="containsText"
        text="urgent"
        fillColor="#FCE4D6"
      />
      <ConditionalFormat
        range="A2:A5"
        type="text"
        operator="beginsWith"
        text="Done"
        fontColor="#548235"
      />
      <Row><Cell>Status</Cell></Row>
      <Row><Cell>urgent: review</Cell></Row>
      <Row><Cell>Done</Cell></Row>
      <Row><Cell>In progress</Cell></Row>
      <Row><Cell>urgent: publish</Cell></Row>
    </Worksheet>
  </Workbook>
);
