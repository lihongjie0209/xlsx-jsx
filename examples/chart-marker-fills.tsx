import { Workbook, Worksheet, Row, Cell, Chart, Series } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Row><Cell>Category</Cell><Cell>Value</Cell></Row>
    <Row><Cell>One</Cell><Cell value={10} /></Row>
    <Row><Cell>Two</Cell><Cell value={20} /></Row>
    <Row><Cell>Three</Cell><Cell value={30} /></Row>
  </Worksheet>
  <Worksheet name="Chart">
    <Chart type="line" cell="D2">
      <Series
        values="Data!$B$2:$B$4"
        categories="Data!$A$2:$A$4"
        markerType="circle"
        markerSize={8}
        markerFillColor="#4472C4"
        markerFillTransparency={25}
        markerLineColor="#336699"
        markerLineWidth={1.5}
        markerLineDash="dashDot"
        markerLineTransparency={25}
      />
      <Series
        values="Data!$B$2:$B$4"
        categories="Data!$A$2:$A$4"
        markerType="diamond"
        markerSize={9}
        markerFillGradient={{type: "linear", angle: 45, stops: [{color: "#C00000", position: 0}, {color: "#FFF2CC", position: 100}]}}
      />
    </Chart>
  </Worksheet>
</Workbook>;
