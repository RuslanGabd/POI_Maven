import java.io.IOException;


public class Main
{
    public static void main(String[] args)
    {
        try
        {
         //   Write.writeIntoExcel("first.xlsx");
            Read.readFromExcel("test.xlsx");
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }

    }

}
