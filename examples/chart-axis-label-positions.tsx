import { Workbook, Worksheet, Chart, Series } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Sales">
      <Chart
        type="column"
        cell="D2"
        title="Quarterly sales"
        horizontalAxisLabelPosition="high"
        verticalAxisLabelPosition="low"
      >
        <Series
          values="Sales!B2:B5"
          categories="Sales!A2:A5"
          title="Revenue"
        />
      </Chart>
    </Worksheet>
  </Workbook>
);
