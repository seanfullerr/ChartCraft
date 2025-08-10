package chartcraft.charts;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.text.XDDFRunProperties;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.SchemaType;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;

public abstract class Chart {
	
	private String title;
	
	private String[] categories;	// X-Axis
	
	private Double[][] data;		// Series data
	
	 private int startCol = 0;
	 
	 private int startRow = 0;	
	 
	 private int colSpan = 5;
	 
	 private int rowSpan = 10;
	 
	 private String titleFont = "Arial";
	 
	 private String xAxisFont = "Arial";
	 
	 private String yAxisFont = "Arial";
	 
	 private String legendFont = "Arial";
	 
	 private double titleFontSize = 12;
	 
	 private double xAxisFontSize = 12;
	 
	 private double yAxisFontSize = 12;
	 
	 private int xAxisRotation = 0;
	 
	 private final int AXIS_ROTATION_MULTIPLIER = 60000	;	
	 
	 private double legendFontSize = 12;
	 
	 private String[] seriesTitles;
	 
	 private LegendPosition legendPosition = LegendPosition.RIGHT;
	 
	 private boolean displayDataLabels = false;
	 
	 private List<int[]> rgb = new ArrayList<>(Arrays.asList((new int[] {0,0,0})));
	 
	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String[] getCategories() {
		return categories;
	}

	public void setCategories(String[] categories) {
		this.categories = categories;
	}

	public Double[][] getData() {
		return data;
	}

	public void setData(Double[][] data) {
		this.data = data;
	}
	
	
	
	
	

	public int getStartCol() {
		return startCol;
	}



	public void setStartCol(int startCol) {
		this.startCol = startCol;
	}



	public int getStartRow() {
		return startRow;
	}



	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}



	public int getColSpan() {
		return colSpan;
	}



	public void setColSpan(int colSpan) {
		this.colSpan = colSpan;
	}



	public int getRowSpan() {
		return rowSpan;
	}



	public void setRowSpan(int rowSpan) {
		this.rowSpan = rowSpan;
	}
	
	

	public List<int[]> getRgb() {
		return rgb;
	}

	public void setRgb(List<int[]> rgb) {
		this.rgb = rgb;
	}

	/**
	 * Starting position of the chart - the col/row the chart's left top corner is in. Must be used prior to creating chart.
	 * 
	 * @param col column chart originates from
	 * @param row row chart originates from
	 */
	public void setPosition(int col, int row) {
		this.startCol = col-1;
	    this.startRow = row-1;
	}
	
	/**
	 * Determines width (columns) and height (rows). Must be used prior to creating chart.
	 * 
	 * @param colsWide width of the chart, total columns
	 * @param rowsTall height of the chart, total rows
	 */
	public void setSpan(int colsWide, int rowsTall) {
		this.colSpan = colsWide;
	    this.rowSpan = rowsTall;
	}
	
	// TODO: SERIES
    /**
     * Sets the bar color for a given series index using RGB values.
     * @param seriesIndex index of the series (0-based)
     * @param r red (0-255)
     * @param g green (0-255)
     * @param b blue (0-255)
     */
    public void setBarColor(int seriesIndex, int[] rgb, XDDFChartData xddfData) {
    	
    	//this.rgb = rgb;
    	
        XDDFChartData.Series series = xddfData.getSeries(seriesIndex);

        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(
            XDDFColor.from(new byte[] {(byte) rgb[0], (byte) rgb[1], (byte) rgb[2]})
        );
        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series.setShapeProperties(properties);
    }
    
    protected XSSFChart setupBlankChart(XSSFSheet sheet) {
    	
    	 XSSFDrawing drawing = sheet.createDrawingPatriarch();
    	 
    	 XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, getStartCol(), getStartRow(),
	              getStartCol() + getColSpan(), getStartRow() + getRowSpan());
	      
    	 
    	 XSSFChart chart = drawing.createChart(anchor);
    			 
    	 chart.setTitleText(getTitle());
	      
	      XDDFChartLegend legend = chart.getOrAddLegend();
	      legend.setPosition(getLegendPosition());
	      
		return chart;
    		
    }
    
    // TODO: SERIES
    /**
     * Sets the bar color for a given series index using RGB values.
     * @param seriesIndex index of the series (0-based)
     * @param r red (0-255)
     * @param g green (0-255)
     * @param b blue (0-255)
     */
    public void setBarColor(List<int[]> rgb) {
    	
    	this.rgb = rgb;
    	
    }
    
    protected void updateFonts(XSSFChart chart) {
    	
    	
    		updateTitleFont(chart);	
    	
    	if ( !getChartType(chart).equals(ChartTypes.PIE) ) {
    		
    		updateAxisFont(chart);
    		
    		
    	}
    	
    	updateLegendFont(chart);
    	
    }
    
    private void updateLegendFont(XSSFChart chart) {
    	
    	  CTLegend ctLegend = chart.getCTChart().getLegend();

	    	
	    	 ctLegend.addNewTxPr();
	    	    ctLegend.getTxPr().addNewBodyPr();
	    	    ctLegend.getTxPr().addNewLstStyle(); // font size in hundreds format (6*100)
	    	    ctLegend.getTxPr().addNewP().addNewPPr().addNewDefRPr().setSz((int)(getLegendFontSize()*100));
	    	    ctLegend.getTxPr().addNewP().addNewPPr().addNewDefRPr().addNewLatin().setTypeface(getLegendFont());
    	
    }
    
   
    private void updateAxisFont(XSSFChart chart) {
    	
    	 CTValAx valAx = ((XSSFChart) chart)
	    	        .getCTChart()
	    	        .getPlotArea()
	    	        .getValAxArray(0);

	    	CTTextBody txPr = valAx.isSetTxPr() ? valAx.getTxPr() : valAx.addNewTxPr();

	    	// Ensure correct order
	    	txPr.addNewBodyPr();
	    	if (!txPr.isSetLstStyle()) txPr.addNewLstStyle();

	    	// Ensure paragraph exists
	    	CTTextParagraph p = (txPr.sizeOfPArray() > 0) ? txPr.getPArray(0) : txPr.addNewP();

	    	// Ensure run properties exist without changing the text
	    	CTTextCharacterProperties rPr = p.isSetPPr() && p.getPPr().isSetDefRPr()
	    	        ? p.getPPr().getDefRPr()
	    	        : (p.isSetPPr() ? p.getPPr().addNewDefRPr() : p.addNewPPr().addNewDefRPr());

	    	// Set font
	    	CTTextFont latin = rPr.isSetLatin() ? rPr.getLatin() : rPr.addNewLatin();
	    	latin.setTypeface(getyAxisFont()); 

	    	//rPr.setB(true);
	    	rPr.setSz((int) (getyAxisFontSize()*100));
	    	

	    	 CTCatAx catAx = ((XSSFChart) chart)
		    	        .getCTChart()
		    	        .getPlotArea()
		    	        .getCatAxArray(0);
	    	 
		    CTTextBody catAxTxPr = catAx.isSetTxPr() ? catAx.getTxPr() : catAx.addNewTxPr();

		    // Ensure correct order
		    catAxTxPr.addNewBodyPr();
	    	if (!catAxTxPr.isSetLstStyle()) catAxTxPr.addNewLstStyle();

	    	// Ensure paragraph exists
	    	CTTextParagraph catP = (catAxTxPr.sizeOfPArray() > 0) ? catAxTxPr.getPArray(0) : catAxTxPr.addNewP();

	    	// Ensure run properties exist without changing the text
	    	CTTextCharacterProperties catRPr = catP.isSetPPr() && catP.getPPr().isSetDefRPr()
	    	        ? catP.getPPr().getDefRPr()
	    	        : (catP.isSetPPr() ? catP.getPPr().addNewDefRPr() : catP.addNewPPr().addNewDefRPr());

	    	// Set font
	    	CTTextFont catLatin = catRPr.isSetLatin() ? catRPr.getLatin() : catRPr.addNewLatin();
	    	catLatin.setTypeface(getxAxisFont()); 
	    	
	    	catRPr.setSz((int) (getxAxisFontSize()*100));
	    	
	    	catAxTxPr.getBodyPr().setRot(getxAxisRotation());
    }
   
    private void updateTitleFont(XSSFChart chart) {
    	
    	 //set "the title overlays the plot area" to false explicitly
	      ((XSSFChart)chart).getCTChart().getTitle().addNewOverlay().setVal(false);

	      //set font style for title - low level
	      //add run properties to title's first paragraph and first text run. Set bold.
	      //((XSSFChart)chart).getCTChart().getTitle().getTx().getRich().getPArray(0).getRArray(0).getRPr().setB(true);
	      //set italic
	      //((XSSFChart)chart).getCTChart().getTitle().getTx().getRich().getPArray(0).getRArray(0).getRPr().setI(true);
	      //set font size 20pt
	      ((XSSFChart)chart).getCTChart().getTitle().getTx().getRich().getPArray(0).getRArray(0).getRPr().setSz( (int) (getTitleFontSize()*100) );
	      //add type face for latin and complex script characters
	      ((XSSFChart)chart).getCTChart().getTitle().getTx().getRich().getPArray(0).getRArray(0).getRPr().addNewLatin().setTypeface(getTitleFont());
	      ((XSSFChart)chart).getCTChart().getTitle().getTx().getRich().getPArray(0).getRArray(0).getRPr().addNewCs().setTypeface(getTitleFont());
	      
	      	
    }
    
    private ChartTypes getChartType(XSSFChart chart) {
    		
    	   CTChart ctChart = chart.getCTChart();
    	    CTPlotArea plotArea = ctChart.getPlotArea();

    	    if (plotArea.getPieChartList().size() > 0) {
    	        return ChartTypes.PIE;
    	    } else if (plotArea.getBarChartList().size() > 0) {
    	        return ChartTypes.BAR;
    	    } else if (plotArea.getLineChartList().size() > 0) {
    	        return ChartTypes.LINE;
    	    } else if (plotArea.getScatterChartList().size() > 0) {
    	        return ChartTypes.SCATTER;
    	    }
    	    // Add more checks if needed

    	    return null;
    	
    }

	public abstract void createChart(XSSFWorkbook workbook, XSSFSheet sheet);

	public String getTitleFont() {
		return titleFont;
	}

	public void setTitleFont(String titleFont) {
		this.titleFont = titleFont;
	}

	public String getxAxisFont() {
		return xAxisFont;
	}

	public void setxAxisFont(String xAxisFont) {
		this.xAxisFont = xAxisFont;
	}

	public String getyAxisFont() {
		return yAxisFont;
	}

	public void setyAxisFont(String yAxisFont) {
		this.yAxisFont = yAxisFont;
	}

	public String getLegendFont() {
		return legendFont;
	}

	public void setLegendFont(String legendFont) {
		this.legendFont = legendFont;
	}

	public LegendPosition getLegendPosition() {
		return legendPosition;
	}

	public void setLegendPosition(LegendPosition legendPosition) {
		this.legendPosition = legendPosition;
	}

	public double getTitleFontSize() {
		return titleFontSize;
	}

	public void setTitleFontSize(double titleFontSize) {
		this.titleFontSize = titleFontSize;
	}

	public double getxAxisFontSize() {
		return xAxisFontSize;
	}

	public void setxAxisFontSize(double xAxisFontSize) {
		this.xAxisFontSize = xAxisFontSize;
	}

	public double getyAxisFontSize() {
		return yAxisFontSize;
	}

		public void setyAxisFontSize(double yAxisFontSize) {
			this.yAxisFontSize = yAxisFontSize;
	}

	public double getLegendFontSize() {
		return legendFontSize;
	}

	public void setLegendFontSize(double legendFontSize) {
		this.legendFontSize = legendFontSize;
	}

	public String[] getSeriesTitles() {
		return seriesTitles;
	}

	public void setSeriesTitles(String[] seriesTitles) {
		this.seriesTitles = seriesTitles;
	}

	public boolean isDisplayDataLabels() {
		return displayDataLabels;
	}

	public void setDisplayDataLabels(boolean displayDataLabels) {
		this.displayDataLabels = displayDataLabels;
	}

	public int getxAxisRotation() {
		return xAxisRotation*AXIS_ROTATION_MULTIPLIER;
	}

	public void setxAxisRotation(int xAxisRotation) {
		this.xAxisRotation = xAxisRotation;
	}
	
	

}
