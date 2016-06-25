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
		frame.setBounds((int) (Toolkit.getDefaultToolkit().getScreenSize().getWidth() - 200), 0, 200, 120);

		// �E�B���h�E��\������
		frame.setVisible(true);
	}

	static class MyFrame extends JFrame {

		/**
		 * 
		 */
		String outputFilePath = "out.docx";
		XWPFDocument document = null;
		XWPFParagraph paragraph;
		XWPFRun run;
		FileOutputStream fout = null;

		Dimension SSize = null;
		int imgnum = 1;
		ArrayList<String> a = new ArrayList<String>();
		ArrayList<Image> imgs = new ArrayList<Image>();
		private static final long serialVersionUID = 1L;
		private int num = 100;

		public int getNum() {
			return num;
		}

		MyFrame(String title) {

			setTitle(title);

			JButton btn_cap = new JButton("cap");
			JButton btn_write = new JButton("write");
			JTextField text1 = new JTextField("", 100);
			JLabel label = new JLabel(new ImageIcon("C:/Users/y-kikuya/Pictures/IMG_0" + getNum() + ".jpg"));

			BoxLayout boxlayout = new BoxLayout(this.getContentPane(), BoxLayout.Y_AXIS);
			this.setLayout(boxlayout);
			this.add(btn_cap);
			this.add(btn_write);
			this.add(text1);
			getContentPane().add(label);

			//�������݃{�^���������ꂽ��
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
							paragraph = document.createParagraph();
							paragraph = document.createParagraph();
							InputStream pic = new FileInputStream("./ShrimpMaker_tmp/" + cnt + ".jpg");
							run.addPicture(pic, Document.PICTURE_TYPE_JPEG, "", Units.toEMU(400),
									Units.toEMU(400 * SSize.getHeight() / SSize.getWidth()));
						}
						fout = new FileOutputStream(outputFilePath);
						document.write(fout);
						System.out.println("�G�r�f���X��������");
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

			//�L���v�`���{�^���������ꂽ��
			btn_cap.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent e) {
					// �摜����荞��ŃA�C�R���ɂ���
					Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
					SSize = screenSize;
					Robot r2d2;
					String text = text1.getText();
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
