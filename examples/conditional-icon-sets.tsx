import { Workbook, Worksheet, Row, Cell, ConditionalFormat } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Scores">
    <ConditionalFormat
      range="A2:A5"
      type="iconSet"
      iconType="threeArrows"
      reverse={false}
      showIconsOnly={false}
    />
    <ConditionalFormat range="B2:B5" type="iconSet" iconType="threeStars" />
    <ConditionalFormat
      range="C2:C5"
      type="iconSet"
      iconType="threeArrows"
      icons={[
        { type: "percent", value: 0 },
        { type: "percentile", value: 50, greaterThan: true, iconType: "fourHistograms", iconIndex: 2 },
        { type: "formula", value: "MAX(C2:C5)", noIcon: true },
      ]}
    />
    <Row>
      <Cell>Score</Cell>
    </Row>
    <Row><Cell>92</Cell></Row>
    <Row><Cell>76</Cell></Row>
    <Row><Cell>48</Cell></Row>
    <Row><Cell>20</Cell></Row>
  </Worksheet>
</Workbook>;
