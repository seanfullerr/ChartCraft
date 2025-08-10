package chartcraft;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import chartcraft.charts.BarChart;
import chartcraft.charts.Chart;
import chartcraft.charts.PieChart;

public class ChartCraft {
	
	static String[] daysOfWeek = new String[]{"Monday", "Tuesday", "Wednesday","Thursday","Friday"};
	
	static String[] seriesTitles = new String[] {"Dogs","Cats"};
	
	static Double[][] totalDogs = new Double[][] {{11.0,4.0,9.0,18.0,6.0},{4.0,5.0,1.0,7.0,4.0}};
	
	static int[] pink = new int[] {255,209,220};
	
	static int[] green = new int[] {190,229,176};



	public static void main(String[] args) throws IOException {
		
		List<int[]> seriesColours = Arrays.asList(pink,green);

		BarChart bc = setupBarChart(new int[] {8,14}, new int[] {3,4}, seriesColours);
			
		BarChart bcTwo = setupBarChart(new int[] {10,16}, new int[] {3,20}, seriesColours);
		
		FileOutputStream f = new FileOutputStream("C:\\Test\\test.xlsx");
		
		XSSFWorkbook wb = new XSSFWorkbook();
		
		wb.createSheet("Test");
		
		bc.createChart(wb, wb.getSheet("Test"));
		
		bcTwo.createChart(wb, wb.getSheet("Test"));
		
		PieChart pc = setupPieChart(new int[] {8,14}, new int[] {3,40});
		
		pc.createChart(wb, wb.getSheet("Test"));
		

		wb.write(f);
		
		f.close();

	}
	
	protected static BarChart setupBarChart(int[] span, int[] pos, List<int[]> colour) {
		
		BarChart bc = new BarChart();
		
		bc.setTitle("Animals petted");
		
		bc.setCategories(daysOfWeek);
		
		bc.setData(totalDogs);
		
		bc.setSpan(span[0], span[1]);
		
		bc.setPosition(pos[0], pos[1]);
		
		bc.setRgb(colour);
		
		bc.setLegendPosition(LegendPosition.RIGHT);
		
		bc.setSeriesTitles(seriesTitles);
		
		bc.setDisplayDataLabels(true);
		
		bc.setxAxisRotation(-	90);
		
		setFonts(bc);
		
		return bc;
		
	}
	
	protected static PieChart setupPieChart(int[] span, int[] pos) {
		
		PieChart pc = new PieChart();
		
		pc.setTitle("Test");
		
		pc.setCategories(daysOfWeek);
		
		pc.setTitleFont("Century Gothic");
		
		pc.setTitleFontSize(11.0);
		
		pc.setLegendFont("Century Gothic");
		
		pc.setLegendFontSize(9.0);
		
		pc.setSpan(span[0], span[1]);
		
		pc.setPosition(pos[0], pos[1]);
		
		return pc;
		
	}
	
	private static void setFonts(Chart c) {
		
		c.setTitleFont("Century Gothic");
		
		c.setTitleFontSize(11.5);
		
		c.setyAxisFont("Century Gothic");
		
		c.setyAxisFontSize(10);
		
		c.setxAxisFont("Century Gothic");
		
		c.setxAxisFontSize(10);
		
		c.setLegendFont("Century Gothic");
		
		c.setLegendFontSize(9.5);
		
	}

}
