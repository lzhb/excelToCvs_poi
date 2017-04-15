package ling;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.Iterator;

/**
 * linglign
 */
public class Ling
{
    //纪录 “车站 ”格子在sheet中的row,column
    static ArrayList<int[]> basePoints = new ArrayList<int[]>();
    static HSSFSheet sheet;

    /**
     * @param args
     * @throws IOException
     */
    public static void main(String[] args) throws IOException
    {
        // 创建 Excel 文件的输入流对象
        FileInputStream excelFileInputStream = new FileInputStream("E:/lingling.xls");
        // XSSFWorkbook 就代表一个 Excel 文件
        // 创建其对象，就打开这个 Excel 文件
        HSSFWorkbook workbook = new HSSFWorkbook(excelFileInputStream);
        // 输入流使用后，及时关闭！这是文件流操作中极好的一个习惯！
        excelFileInputStream.close();
        // 我们通过 getSheetAt(0) 指定表格索引来获取对应表格
        // 注意表格索引从 0 开始！
//		for(int i=0;i<workbook.getNumberOfSheets();i++){
        for (int i = 0; i < workbook.getNumberOfSheets(); i++)
        {
            sheet = workbook.getSheetAt(i);
//			String aa=sheet.getRow(8).getCell(1).getStringCellValue();
//			String bb=sheet.getRow(8).getCell(5).getStringCellValue();
//		    System.out.println(aa.equals(bb));
//		    System.out.println("===");

            for (Row hssfRow : sheet)
            {
                for (Cell iCell : hssfRow)
                {
                    if (iCell.getStringCellValue().equals("车站"))
                    {
                        //第三个 数据 为0 表示 该瞄点相关的小表还未被访问。 若为1，则已访问。  方便标记处special_2_basePoint
                        int[] point = new int[]{hssfRow.getRowNum(), iCell.getColumnIndex(), 0};
                        basePoints.add(point);
                    }
                }
            }

//			int [] p=basePoints.get(0); //第一个sheet第一张小表的 瞄点  // 82 第8行第2个。!! 0-based
//			System.out.println(p[0]+""+p[1]);

            FileOutputStream bos = new FileOutputStream("E:/output_sheet" + i + ".txt");
            System.setOut(new PrintStream(bos));


            //遍历 basepoints瞄点，并把结果写入文件中
            for (int[] p : basePoints)
            {
                if (p[2] == 0)
                {
                    //读取左边的车列
                    printL(sheet, p[0], p[1]);
                    //读取右边的车列
                    printR(sheet, p[0], p[1]);
                    p[2] = 1;

                }
            }
            bos.close();
        }


    }

    //	车次,站序,车站,到达时刻,离开时刻
//	D6091,1,即墨北,,6:50
//	D6091,2,莱阳,7:17,7:19
//	D6091,3,桃村北,7:40,7:41
//	D6091,4,烟台,8:10,,
    private static void printR(HSSFSheet sheet, int rowNum, int colIndex)
    {
        String tName = sheet.getRow(rowNum).getCell(colIndex + 1).getStringCellValue(); //G2
        if (tName.equals(""))
            return;//特例 23sheet中 没有 右边的列车
        String startArrival = sheet.getRow(rowNum - 5).getCell(colIndex + 1).getStringCellValue();
        String endArrival = sheet.getRow(rowNum - 3).getCell(colIndex + 1).getStringCellValue();

        System.out.println("车次,站序,车站,到达时刻,离开时刻");

        int i = 0;//站序
        int row = rowNum;//行
        String cArrival = "";

        if (isdivided)
        {
            row = special_2_endPoint[0];
            colIndex = special_2_endPoint[1];
        } else
        {
            //正常情况下， 不是两张小表
            while (cArrival != startArrival)
            {
                row++;
                cArrival = sheet.getRow(row).getCell(colIndex).getStringCellValue();
            }
            row++;
        }

        do
        {
            if (isdivided && (row == special_2_basePoint[0]))
            {
                //调到 上表 的末尾
                row = special_1_endPoint[0];
                colIndex = special_1_endPoint[1];

                isdivided = false;
            }
            String aTime = sheet.getRow(row).getCell(colIndex + 1).getStringCellValue();
            cArrival = sheet.getRow(--row).getCell(colIndex).getStringCellValue();
            String bTime = sheet.getRow(row--).getCell(colIndex + 1).getStringCellValue();
            if (aTime.equals("") && bTime.equals(""))
            {
                //若某个站点 不停留，就不输出 该条信息
                continue;
            }
            i++;
            if (bTime.equals("--")) bTime = ""; //不能用 ==
            if (bTime.length() == 2)
            {
                if (aTime.length() == 5)
                    bTime = aTime.substring(0, 3) + bTime;
                if (aTime.length() == 4)
                    bTime = aTime.substring(0, 2) + bTime;
            }
            System.out.println(tName + "," + i + "," + cArrival + "," + aTime + "," + bTime);
        } while (cArrival != endArrival);
    }

    /**
     * @param s
     * @param row
     * @param col
     * @return
     */
    //寻找两张表中的第二张
    public static int[] searchAnother(String s, int row, int col)
    {
        //row, col 是 车次旁的 “车站” 的位置
        for (Row hssfRow : sheet)
        {
            for (Cell iCell : hssfRow)
            {
                if (iCell.getStringCellValue().equals(s))
                {
                    if ((hssfRow.getRowNum() != row) || ((iCell.getColumnIndex() + 1) != col))
                    { //要+1
                        //修正到 "车站" 瞄点的位置上
                        int[] point = new int[]{hssfRow.getRowNum(), iCell.getColumnIndex() + 1};
                        return point;
                    }
                }
            }
        }
        return null;
    }

    static int[] special_2_basePoint;//第二张表的瞄点
    static int[] special_2_endPoint;//第二张表的表尾瞄点
    static int[] special_1_endPoint;//第一张表的表尾瞄点
    static boolean rflag = false;
    static boolean isdivided;//是否被分成两张小表


    public static void printL(HSSFSheet sheet, int rowNum, int colIndex)
    {
        String tName = sheet.getRow(rowNum).getCell(colIndex - 1).getStringCellValue(); //G1
        if (tName.equals(""))
            return;//特例 22sheet中 没有 左边的列车

//		System.out.println("***  tName -- "+tName);
        String endArrival = sheet.getRow(rowNum - 3).getCell(colIndex - 1).getStringCellValue(); //注意是合并格，得-3
        String startArrival = sheet.getRow(rowNum - 5).getCell(colIndex - 1).getStringCellValue();
        String checkArrival = sheet.getRow(rowNum + 1).getCell(colIndex).getStringCellValue();
        boolean flag_shang = false;
        if (startArrival == checkArrival)
            flag_shang = true;//该表是上部分的表 或 完整的表
        String cArrival;

        //检查小表是否完整
        int endrow = rowNum + 1;//行
        while ((sheet.getRow(endrow) != null) && (sheet.getRow(endrow).getCell(colIndex) != null) && (sheet.getRow(endrow).getCell(colIndex).getStringCellValue().length() > 0))
        {
            endrow = endrow + 2;
        }
        endrow--;   //62
//		System.out.println("***  endrow -- "+endrow);

//		String endTest=sheet.getRow(endrow-1).getCell(colIndex).getStringCellValue();

//		System.out.println("  endArrival:"+endArrival);
//		System.out.println("  endTest:"+endTest);
        isdivided = false;//是否被分成两张小表 , 重置为false

        special_2_basePoint = searchAnother(tName, rowNum, colIndex);
        if (special_2_basePoint != null)
        {
            //不是一个完整的小表
            isdivided = true;
            if (flag_shang)
            {
                //若是 上表
                //找到第二张表的瞄点， 首先需要保证 该表 是上部分的表


                //在basePoints中 将 第二张表的瞄点 标记处，掠过，不访问它
                Iterator<int[]> iter = basePoints.iterator();
                while (iter.hasNext())
                {
                    int[] p = (int[]) iter.next();
                    if ((p[0] == special_2_basePoint[0]) && (p[1] == special_2_basePoint[1]))
                    {
                        p[2] = 1;
                    }
                }
                //设一个rflag，标识该表有没有右边的列车
                //若有，rflag=ture. 并把第二张表的瞄点、第一张表的表尾瞄点  纪录为static 常量，方便printR 使用
                String tName_r = sheet.getRow(rowNum).getCell(colIndex + 1).getStringCellValue(); //G2
                if (tName_r.equals(""))
                {
                    rflag = false;// 没有 右边的列车
                } else
                {
                    rflag = true;
                    special_1_endPoint = new int[]{endrow, colIndex};
//					System.out.println("-----special_1_endPoint :"+special_1_endPoint[0]+"*"+special_1_endPoint[1]);

                }

            } else
            {
                //若是 一开始就扫描到 下表
                //***
            }

        }
//		System.out.println("  isdivided:"+isdivided);
        System.out.println("车次,站序,车站,到达时刻,离开时刻");
        int i = 0;//站序
        int row = rowNum;//行
        do
        {
            if ((row == endrow) && (isdivided == true))
            {

//				System.out.println("----------***转到表二***");
                //将 row colIndex修改 到第二张表上
                row = special_2_basePoint[0];
                colIndex = special_2_basePoint[1];
            }
            cArrival = sheet.getRow(++row).getCell(colIndex).getStringCellValue();
            String aTime = sheet.getRow(row).getCell(colIndex - 1).getStringCellValue();
            String bTime = sheet.getRow(++row).getCell(colIndex - 1).getStringCellValue();
            if (aTime.equals("") && bTime.equals(""))
            {
                //若某个站点 不停留，就不输出 该条信息
                continue;
            }
            i++;
            if (bTime.equals("--")) bTime = ""; //不能用 ==
            if (bTime.length() == 2)
            {
                if (aTime.length() == 5)
                    bTime = aTime.substring(0, 3) + bTime;
                if (aTime.length() == 4)
                    bTime = aTime.substring(0, 2) + bTime;
            }
            System.out.println(tName + "," + i + "," + cArrival + "," + aTime + "," + bTime);
        } while (cArrival != endArrival);
        //若rflag=ture. 并把第二张表的表尾瞄点 纪录为static 常量
        if (isdivided && rflag)
        {
            special_2_endPoint = new int[]{row, colIndex};
//			System.out.println("-----special_2_endPoint :"+special_2_endPoint[0]+"*"+special_2_endPoint[1]);

        }

    }
}
