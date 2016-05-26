package data;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
public class text3 {

	public static void main(String[] args) throws IOException {


		Scanner sc = new Scanner(System.in);
		System.out.println("Wordファイルが存在するフォルダのフルパスを入力：");
		String file =sc.next();
		File filename = new File(file);
		File[] files1 = filename.listFiles();
		System.out.println(files1.length);
		for(int i = 0;i < files1.length;i++){
			FileInputStream fis = new FileInputStream(files1[i]);
			String para = "";
			int z = 0;
			if(files1[i].getName().endsWith(".doc")){
				HWPFDocument doc = new HWPFDocument(fis);
				WordExtractor we = new WordExtractor(doc);
				para = we.getText();
				z = 2;
			}else if(files1[i].getName().endsWith(".docx")){
				XWPFDocument doc = new XWPFDocument(fis);
				XWPFWordExtractor we = new XWPFWordExtractor(doc);
				para = we.getText();
				z = 1;
			}else{
				continue;
			}


			int sum = para.length()  ;


			System.out.println(files1[i] + "の文字数：" +  (para.length() - z)); // 文字数 + 段落 - 1
			fis.close();
		}
	}
}



