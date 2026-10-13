import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet
      name="Regions"
      autoFilter="A1:A4"
      autoFilterColumns={[
        {
          column: 0,
          customFilters: {
            logic: "or",
            conditions: [
              { operator: "beginsWith", value: "North" },
              { operator: "doesNotContain", value: "Test" },
            ],
          },
        },
      ]}
    >
      <Row><Cell>Region</Cell></Row>
      <Row><Cell>North</Cell></Row>
      <Row><Cell>South</Cell></Row>
      <Row><Cell>Test Area</Cell></Row>
    </Worksheet>
  </Workbook>
);
