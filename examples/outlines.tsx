import { Workbook, Worksheet, Column, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Outline">
      <Column min={2} max={3} outlineLevel={1} />
      <Row><Cell>Region total</Cell><Cell>1200</Cell></Row>
      <Row outlineLevel={1}><Cell>North</Cell><Cell>700</Cell></Row>
      <Row outlineLevel={2}><Cell>North detail</Cell><Cell>300</Cell></Row>
      <Row outlineLevel={1}><Cell>South</Cell><Cell>500</Cell></Row>
    </Worksheet>
  </Workbook>
);
