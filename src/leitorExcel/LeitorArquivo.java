package leitorExcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.collections4.IteratorUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.Cleanup;

public class LeitorArquivo {

	public ArrayList<Cliente> leitorNomes() throws IOException {

		ArrayList<Cliente> listaClientes = new ArrayList<>();

		@Cleanup
		FileInputStream file = new FileInputStream("src/leitorExcel/clientes.xlsx");

		Workbook workbook = new XSSFWorkbook(file);

		Sheet sheet = workbook.getSheetAt(0);

		List<Row> rows = (List<Row>) toList(sheet.iterator());

		DataFormatter formatter = new DataFormatter();

		rows.forEach(row -> {

			if (row.getRowNum() == 0) {

			} else {

				List<Cell> cells = (List<Cell>) toList(row.iterator());

				String codigo;
				String razao;
				String fantasia;
				String endereco;

				if (cells.get(0).getCellType() == CellType.NUMERIC) {
					codigo = String.valueOf(cells.get(0).getNumericCellValue());
				} else {
					codigo = cells.get(0).getStringCellValue();
				}

				if (cells.get(1).getCellType() == CellType.NUMERIC) {
					razao = String.valueOf(cells.get(0).getNumericCellValue());
				} else {
					razao = cells.get(1).getStringCellValue();
				}

				if (cells.get(2).getCellType() == CellType.NUMERIC) {
					fantasia = String.valueOf(cells.get(0).getNumericCellValue());
				} else {
					fantasia = cells.get(2).getStringCellValue();
				}

				if (cells.get(3).getCellType() == CellType.NUMERIC) {
					endereco = String.valueOf(cells.get(0).getNumericCellValue());
				} else {
					endereco = cells.get(3).getStringCellValue();
				}

				Cliente cliente = new Cliente();
				cliente.setCodigo(codigo);
				cliente.setRazaoSocial(razao);
				cliente.setFantasia(fantasia);
				cliente.setEndereco(endereco);

				listaClientes.add(cliente);

			}

		});

		return listaClientes;

	}

	public List<?> toList(Iterator<?> iterator) {
		return IteratorUtils.toList(iterator);
	}

}
