package aaa;

import java.awt.AWTException;
import java.awt.Dimension;
import java.awt.Image;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import javax.imageio.ImageIO;
import javax.swing.BoxLayout;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;

import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ImgConv {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		File newdir = new File("./ShrimpMaker_tmp");
		if (!newdir.exists()) {
			newdir.mkdir();
		}
		MyFrame frame = new MyFrame("ShrimpMaker");
		;
		frame.setBounds((int) (Toolkit.getDefaultToolkit().getScreenSize().getWidth() - 200), 0, 200, 150);

		// ウィンドウを表示する
		frame.setVisible(true);
	}

	static class MyFrame extends JFrame {

		/**
		 * 
		 */
		String outputFileExtension = ".docx";
		XWPFDocument document = null;
		XWPFParagraph paragraph;
		XWPFRun run;
		FileOutputStream fout = null;

		Dimension SSize = null;
		int imgnum = 1;
		ArrayList<String> a = new ArrayList<String>();
		ArrayList<Image> imgs = new ArrayList<Image>();
		private static final long serialVersionUID = 1L;
		
		MyFrame(String title) {

			setTitle(title);

			JButton btn_cap = new JButton("cap");
			JButton btn_write = new JButton("write");
			JTextField text1 = new JTextField("", 100);
			JTextField filename = new JTextField("", 100);
			JLabel label = new JLabel(new ImageIcon());

			BoxLayout boxlayout = new BoxLayout(this.getContentPane(), BoxLayout.Y_AXIS);
			this.setLayout(boxlayout);
			this.add(text1);
			this.add(btn_cap);
			this.add(filename);
			this.add(btn_write);
			getContentPane().add(label);

			//書き込みボタンが押されたら
			btn_write.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent e) {
					try {
						document = new XWPFDocument();

						int cnt = 0;
						for (String i : a) {
							cnt++;

							paragraph = document.createParagraph();
							paragraph.setAlignment(ParagraphAlignment.LEFT);
							run = paragraph.createRun();

							run.setText(cnt + "." + i);
							run.addCarriageReturn();  //改行
							paragraph = document.createParagraph();
							InputStream pic = new FileInputStream("./ShrimpMaker_tmp/" + cnt + ".jpg");
							run.addPicture(pic, Document.PICTURE_TYPE_JPEG, "", Units.toEMU(400),
									Units.toEMU(400 * SSize.getHeight() / SSize.getWidth()));
						}
						String outputFilePath;
						if(filename.getText().equals("") || filename.getText().isEmpty()) {
							outputFilePath = "output.docx";
						} else {
							outputFilePath = filename.getText() + outputFileExtension;
						}
						fout = new FileOutputStream(outputFilePath);
						document.write(fout);
						//arraylistに入っている出力済み情報をクリア
						imgnum = 1;
						a.clear();
						imgs.clear();
						
						System.out.println("エビデンス生成成功");
					} catch (FileNotFoundException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					} catch (InvalidFormatException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				}
			});

			//キャプチャボタンが押されたら
			btn_cap.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent e) {
					// 画像を取り込んでアイコンにする
					Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
					SSize = screenSize;
					Robot r2d2;
					String text = text1.getText();
					text1.setText("");
					a.add(text);
					try {
						r2d2 = new Robot();
						BufferedImage img = r2d2.createScreenCapture(new Rectangle(screenSize));
						Image display_img = img.getScaledInstance(
								(int) (img.getWidth(label) * 792 / img.getWidth(label)), -1, Image.SCALE_SMOOTH);

						label.setIcon(new ImageIcon(display_img));
						imgs.add(img);
						try {
							ImageIO.write(img, "jpg", new File("./ShrimpMaker_tmp/" + imgnum + ".jpg"));
							imgnum++;
						} catch (Exception e1) {
							e1.printStackTrace();
						}

					} catch (AWTException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				}
			});

			ImageIcon icon = new ImageIcon("./img/evi.png");
			setIconImage(icon.getImage());

			setLocationRelativeTo(null);
			setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		}
	}
}
