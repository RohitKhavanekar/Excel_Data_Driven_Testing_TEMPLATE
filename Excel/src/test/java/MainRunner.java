import java.io.IOException;
import java.lang.reflect.Array;
import java.util.ArrayList;

public class MainRunner {

	public static void main(String[] args) throws IOException 
	{
		// TODO Auto-generated method stub
		ExcelFile e = new ExcelFile();
		ArrayList data = e.getdata("Add Profile");
		
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));
		

	}

}
