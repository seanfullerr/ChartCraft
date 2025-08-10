package chartcraft.charts;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.chart.AxisCrossBetween;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;

public class BarChart extends Chart {

	@Override
	public void createChart(XSSFWorkbook workbook, XSSFSheet sheet) {

		System.out.println("Creating chart "+getTitle());
		

	      XSSFChart chart =  setupBlankChart(sheet);
	    		  
	      // Use a category axis for the bottom axis.
	      XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
	      XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
	      leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
	      leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
	      
	      XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);

	      XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromArray(getCategories());

	      for (int i = 0; i < getData().length; i++) {
	          XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromArray(getData()[i]);

	          XDDFBarChartData.Series barSeries = (XDDFBarChartData.Series) data.addSeries(xs, values);

	          CTSerTx tx = barSeries.getCTBarSer().getTx();
	          if (tx.isSetStrRef()) {
	              tx.unsetStrRef();
	          }

	          barSeries.setTitle(getSeriesTitles()[i], null);
	          
	          if ( isDisplayDataLabels() ) {
	        	  
	        	  barSeries.setShowLeaderLines(true);
	        	 
		          updateDatLabels(chart, i);
	        	  
	          }

	          if (getRgb() != null && getRgb().size() > i && getRgb().get(i) != null) {
	              setBarColor(i, getRgb().get(i), data);
	          }
	      }
	      
	      
          chart.plot(data);

	      //repairing set the kind of bar char, either bar chart or column chart:
	      if (chart.getCTChart().getPlotArea().getBarChartArray(0).getBarDir() == null) 
	       chart.getCTChart().getPlotArea().getBarChartArray(0).addNewBarDir();
	      chart.getCTChart().getPlotArea().getBarChartArray(0).getBarDir().setVal(
	       org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir.COL);

	      
	      //repairing telling the axis Ids in bar chart:
	      if (chart.getCTChart().getPlotArea().getBarChartArray(0).getAxIdList().size() == 0) {
	       chart.getCTChart().getPlotArea().getBarChartArray(0).addNewAxId().setVal(bottomAxis.getId());
	       chart.getCTChart().getPlotArea().getBarChartArray(0).addNewAxId().setVal(leftAxis.getId());
	      }

	      updateFonts(chart);
	}
	
	private void updateDatLabels(XSSFChart chart, int serIter) {
		
		 chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(serIter).getDLbls()
         .addNewDLblPos().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos.OUT_END);
	      chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(serIter).getDLbls().addNewShowVal().setVal(true);
	      chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(serIter).getDLbls().addNewShowLegendKey().setVal(false);
	      chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(serIter).getDLbls().addNewShowCatName().setVal(false);
	      chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(serIter).getDLbls().addNewShowSerName().setVal(false);
		
	}

}
