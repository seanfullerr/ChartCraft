package chartcraft.charts;

import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PieChart extends Chart {

	@Override
	public void createChart(XSSFWorkbook workbook, XSSFSheet sheet) {

		      
		      
	            XSSFChart chart = setupBlankChart(sheet);

	            XDDFChartLegend legend = chart.getOrAddLegend();
	            legend.setPosition(getLegendPosition());

	            XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(getCategories());
	            XDDFNumericalDataSource<Double> val = XDDFDataSourcesFactory.fromArray(new Double[]{170d, 99d, 98d});

	            //XDDFChartData data = new XDDFPieChartData(chart.getCTChart().getPlotArea().addNewPieChart());
	            XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
	            data.setVaryColors(true);
	            XDDFChartData.Series series = data.addSeries(cat, val);
	            series.setTitle("Series", null);
	            chart.plot(data);

		 updateFonts(chart);
			
	}

}
