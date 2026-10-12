import { Workbook, Worksheet, Row, Cell, Chart, Series } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Data">
      <Chart type="line" cell="E2" x2AxisTitle="Secondary period" y2AxisTitle="Margin" x2AxisLabelFontSize={10} x2AxisLabelFontBold={true} y2AxisLabelFontSize={9} y2AxisLabelFontItalic={true} x2AxisMajorTickType="inside" y2AxisMinorTickType="outside" x2AxisLabelPosition="high" y2AxisLabelPosition="low" x2AxisHidden={false} y2AxisHidden={false} secondaryCategoryAxisLabelInterval={2} secondaryCategoryAxisTickInterval={2} secondaryValueAxisDisplayUnit="thousands" secondaryValueAxisDisplayUnitLabel={false} x2AxisCrossing={2} y2AxisCrossing={0.2} x2AxisMinorGridlines={true} y2AxisMajorGridlines={{color: "#A5A5A5", width: 0.75, dash: "dash", transparency: 10}} y2AxisNumberFormat="0.0%" secondaryValueAxisMin={0.1} secondaryValueAxisMax={1} secondaryValueAxisMajorUnit={0.25} secondaryValueAxisMinorUnit={0.05} secondaryValueAxisLogBase={10} secondaryValueAxisReversed={true}>
        <Series values="Data!$C$2:$C$4" categories="Data!$D$2:$D$4" secondaryAxis={true} />
        <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" />
      </Chart>
      <Row>
        <Cell>Quarter</Cell><Cell>Revenue</Cell><Cell>Margin</Cell><Cell>Month</Cell>
      </Row>
      <Row>
        <Cell>Q1</Cell><Cell>125000</Cell><Cell>0.18</Cell><Cell>Jan</Cell>
      </Row>
      <Row>
        <Cell>Q2</Cell><Cell>147000</Cell><Cell>0.21</Cell><Cell>Feb</Cell>
      </Row>
      <Row>
        <Cell>Q3</Cell><Cell>139000</Cell><Cell>0.19</Cell><Cell>Mar</Cell>
      </Row>
    </Worksheet>
  </Workbook>
);
