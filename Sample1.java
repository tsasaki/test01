//package sample;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sample1 {

  public static void main(String[] args) throws IOException {

    // Excel2007�ȍ~�́u.xlsx�v�`���̃t�@�C���̑f���쐬
    Workbook book = new XSSFWorkbook();

    // �V�[�g���u�T���v���v�Ƃ������O�ō쐬
    Sheet sheet = book.createSheet("�T���v��");

    // 1�s�ڍ쐬 ��Excel��A�s�ԍ���1����X�^�[�g���Ă܂����A
    // �\�[�X���ł�0����̃X�^�[�g�ɂȂ��Ă���̂ŗv���ӁI
    Row row = sheet.createRow(0);

    // 1�ڂ̃Z�����쐬 ���s�Ɠ������A0����X�^�[�g
    Cell a1 = row.createCell(0);  // Excel��A�uA1�v�̏ꏊ

    // �l���Z�b�g
    a1.setCellValue("POI�̃e�X�g");

    // �Z���̃X�^�C����ς��Ă݂悤�Ǝv���܂��E�E�E
    CellStyle style =  book.createCellStyle();
    // �t�H���g
    Font font = book.createFont();
    // ROSE�F�ɂ��Ă݂�
    // IndexedColors�ɂ͂����ȐF���ݒ肳��Ă܂�
    font.setColor(IndexedColors.ROSE.getIndex());
    // �Z���ɃZ�b�g�I�I
    style.setFont(font);
    a1.setCellStyle(style);

    // ��������o�͏���
    FileOutputStream out = null;
    try {
	// �o�͐�̃t�@�C�����w��
	out = new FileOutputStream("C:\\temp/Sample1.xlsx");
	// ��L�ō쐬�����u�b�N���o�͐�ɏ�������
	book.write(out);
    
    } catch (FileNotFoundException e) {
	System.out.println(e.getStackTrace());

    } finally {
	// �Ō�͂����ƕ��Ă����܂�
	out.close();
	book.close();
    }
  }
}
