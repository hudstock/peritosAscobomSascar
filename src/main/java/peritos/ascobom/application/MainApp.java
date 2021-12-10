/*
		 * Obsevações Decidi criar 3 métodos diferentes para processar os arquivos
		 * devido grandes diferenças encontradas nos arquivos. Para não ter que alterar
		 * a estrutura dos arquivos, optei por fazer pequenos ajustes duplicando os
		 * métodos. E também por se tratar de um aplicativo que terá a execução para
		 * apenas um trabalho específico. Com o término deste projeto, este código fonte
		 * será praticamente descartado.
		 *
		 * Foi utilizado JDBC e não JPAm visando máximo de performance ao processar os
		 * arquivos, além da não definição inicial das entidades de domínio da
		 * aplicação.
		 */

package peritos.ascobom.application;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.sql.Connection;
import java.sql.Date;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.temporal.ChronoUnit;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mysql.cj.result.LocalDateTimeValueFactory;

import peritos.ascobom.db.ConnectionFactory;
import peritos.ascobom.db.DbUtils;

public class MainApp {

	public static void main(String[] args) throws Exception {

		LocalDateTime inicio;
		LocalDateTime fim;
		Connection conexaoBD = ConnectionFactory.getConnection();

		inicio = LocalDateTime.now();
		
		processar(conexaoBD, "C:\\Dev\\Arquivos\\Tratadas\\ano2010Julho.xlsx");
		
		/*processar(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2008.xlsx");
		processar(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2009.xlsx");
		processar(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2010.xlsx");

		// Método para processar apenas as 2 primeiras abas do arquivo 2011
		processar2011(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2011.xlsx");

		// Uma execução por aba faz processamento final ficar mais rápido.
		// O arquivo 2011 é grande demais.
		processar2011Contrato(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2011.xlsx", "MARÇO 2011");
		processar2011Contrato(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2011.xlsx", "ABRIL 2011");
		processar2011Contrato(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2011.xlsx", "MAIO 2011");
		processar2011Contrato(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2011.xlsx", "JUNHO 2011");
		processar2011Contrato(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2011.xlsx", "JULHO 2011");
		processar2011Contrato(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2011.xlsx",
				"AGOSTO 2011");
		processar2011Contrato(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2011.xlsx",
				"SETEMBRO 2011");
		processar2011Contrato(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2011.xlsx",
				"OUTUBRO 2011");
		processar2011Contrato(conexaoBD, "/home/hud/Desktop/ProjetoAscobom/ArquivosAscobom/ano2011.xlsx",
				"NOVEMBRO 2011");
		*/

		System.out.println("Fim execução");

		fim = LocalDateTime.now();
		System.out.println(inicio);
		System.out.println(fim);

		long minutes = fim.until(inicio, ChronoUnit.MINUTES);
		long seconds = fim.until(inicio, ChronoUnit.SECONDS);

		System.out.println("Tempo execução: " + minutes + " minutos " + seconds + " segundos.");
	}

	public static void processar(Connection conexaoBD, String arquivo) throws Exception {

		XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(arquivo));
		Iterator<Sheet> sheetIterator = wb.sheetIterator();
		CellStyle estiloNegrito = getEstiloNegrito(wb);
		CellStyle estiloSemNegrito = getEstiloSemNegrito(wb);

		String coordenadaValorOperante = "F5";
		String coordenadaValorInoperante = "F6";
		String coordenadaValorNaoAceitoAscobom = "F7";
		// Iterando nas Abas da planilha
		while (sheetIterator.hasNext()) {
			Sheet abaAtual = sheetIterator.next();
			java.util.Date dataAba = DbUtils.geraDataFimAba(abaAtual.getSheetName());
			// DateTimeFormatter formatter =
			// DateTimeFormatter.ofPattern("dd/MM/yyyy",Locale.US);
			Iterator<Row> linhaIterator = abaAtual.iterator();
			double valorOperante = 0;
			double valorInoperante = 0;
			double valorNaoVerificavel = 0;

			inserirStringNaPlanilha(abaAtual, "TOTAL OPERANTE", "E5", estiloNegrito);
			inserirStringNaPlanilha(abaAtual, "TOTAL INOPERANTE", "E6", estiloNegrito);
			inserirStringNaPlanilha(abaAtual, "TOTAL NÃO VERIFICÁVEL", "E7", estiloNegrito);

			// Iterando nas linhas da aba aberta
			while (linhaIterator.hasNext()) {
				Row linha = linhaIterator.next();

				// Linha que cabeçalho deve ser criado
				if (linha.getRowNum() == 8) {
					criarCabecalhoResultadoPlanilha(wb, estiloNegrito, linha);
				}

				// Linha que inicia a ter dados para processamento
				if (linha.getRowNum() > 8) {
					Cell celPlaca = linha.getCell(0);
					if (celPlaca == null) {
						continue;
					}
					celPlaca.setCellType(CellType.STRING);
					String placa = celPlaca.getStringCellValue();

					if (placa != null && placa != "") {
						Cell cellValor = linha.getCell(1);
						String statusOriginal = linha.getCell(2).getStringCellValue();
						double valorProcessamento = 0.0;
						double valorOriginal = 0.0;

						if (cellValor != null) {
							valorOriginal = cellValor.getNumericCellValue();
						}
						if (valorOriginal == 0) {
							valorProcessamento = 60;
						} else {
							valorProcessamento = valorOriginal;
						}
						Date dataFimPlaca = DbUtils.consultarPlaca(conexaoBD, placa);
						String resultado;
						String textoData;
						
						if (dataFimPlaca != null) {
							if (dataAba.after(dataFimPlaca)) {
								resultado = "Inoperante";
								valorInoperante += valorProcessamento;
							} else {
								resultado = "Operante";
								valorOperante += valorProcessamento;
							}
							textoData = formatarData(dataFimPlaca);
						} else {
							resultado = "Não Verificável";
							// valorOperante += valor;
							valorNaoVerificavel += valorProcessamento;
							textoData = "Veículo não cadastrado.";							
						}

						Cell cellResultado = linha.createCell(4);
						cellResultado.setCellValue(resultado);
						cellResultado.setCellStyle(estiloSemNegrito);

						Cell cellData = linha.createCell(5);
						cellData.setCellStyle(estiloSemNegrito);
						cellData.setCellType(CellType.STRING);
						cellData.setCellValue(textoData);

						System.out.println("Placa:" + placa + " -  Resultado:" + resultado);
						System.out.println("Aba em processamento:" + formatarData(dataAba));
						System.out.println("Data fim placa encontrada:" + textoData);
						// System.out.println("Valor total operante:" + valorOperante + "Valor total
						// inoperante:" + valorInoperante);

						DbUtils.inserirRegistroResultado(conexaoBD, dataAba, placa, null, getBigDecimal(valorOriginal),
								getBigDecimal(valorProcessamento), linha.getCell(2).getStringCellValue(), resultado,
								dataFimPlaca);
					}
				}

			}
			inserirStringNaPlanilha(abaAtual, getValorNumericoFormatado(valorOperante), coordenadaValorOperante,
					estiloSemNegrito);
			inserirStringNaPlanilha(abaAtual, getValorNumericoFormatado(valorInoperante), coordenadaValorInoperante,
					estiloSemNegrito);
			inserirStringNaPlanilha(abaAtual, getValorNumericoFormatado(valorNaoVerificavel),
					coordenadaValorNaoAceitoAscobom, estiloSemNegrito);

			double valorCobradoPagoAscobom = getValorNumericoPlanilha(abaAtual, "B4");
			double valorDevidoAscobom = getValorNumericoPlanilha(abaAtual, "E4");
			double valorTotalAscobom = getValorNumericoPlanilha(abaAtual, "C8");

			DbUtils.inserirTotal(conexaoBD, dataAba, getBigDecimal(valorOperante), getBigDecimal(valorInoperante),
					getBigDecimal(valorNaoVerificavel), getBigDecimal(valorCobradoPagoAscobom),
					getBigDecimal(valorDevidoAscobom), getBigDecimal(valorTotalAscobom));
			autoSize(abaAtual);

		}
		FileOutputStream fileOut = new FileOutputStream(arquivo);
		wb.write(fileOut);
		wb.close();
		fileOut.close();
	}

	private static void autoSize(Sheet abaAtual) {
		for (int x = 0; x < 5; x++) {
			abaAtual.autoSizeColumn(x);
		}
	}

	// Este método é para processar apenas o arquivo de 2011, as duas primeiras abas
	public static void processar2011(Connection conexaoBD, String arquivo) throws Exception {

		XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(arquivo));
		Iterator<Sheet> sheetIterator = wb.sheetIterator();
		CellStyle estiloNegrito = getEstiloNegrito(wb);
		CellStyle estiloSemNegrito = getEstiloSemNegrito(wb);

		String coordenadaValorOperante = "F5";
		String coordenadaValorInoperante = "F6";
		String coordenadaValorNaoAceitoAscobom = "F7";
		// Iterando nas Abas da planilha
		while (sheetIterator.hasNext()) {
			Sheet abaAtual = sheetIterator.next();
			if (abaAtual.getSheetName().equals("JANEIRO 2011") || abaAtual.getSheetName().equals("FEVEREIRO 2011")) {
				java.util.Date dataAba = DbUtils.geraDataFimAba(abaAtual.getSheetName());

				// DateTimeFormatter formatter =
				// DateTimeFormatter.ofPattern("dd/MM/yyyy",Locale.US);
				Iterator<Row> linhaIterator = abaAtual.iterator();
				double valorOperante = 0;
				double valorInoperante = 0;
				double valorNaoVerificavel = 0;

				inserirStringNaPlanilha(abaAtual, "TOTAL OPERANTE", "E5", estiloNegrito);
				inserirStringNaPlanilha(abaAtual, "TOTAL INOPERANTE", "E6", estiloNegrito);
				inserirStringNaPlanilha(abaAtual, "TOTAL NÃO VERIFICÁVEL", "E7", estiloNegrito);

				// Iterando nas linhas da aba aberta
				while (linhaIterator.hasNext()) {
					Row linha = linhaIterator.next();

					// Linha que cabeçalho deve ser criado
					if (linha.getRowNum() == 9) {
						criarCabecalhoResultadoPlanilha(wb, estiloNegrito, linha);
					}

					// Linha que inicia a ter dados para processamento
					if (linha.getRowNum() > 9) {
						Cell celPlaca = linha.getCell(0);
						if (celPlaca == null) {
							continue;
						}
						celPlaca.setCellType(CellType.STRING);
						String placa = celPlaca.getStringCellValue();

						if (placa != null && placa != "") {
							Cell cellValor = linha.getCell(1);
							double valorOriginal = 0.0;
							double valorProcessamento = 0.0;
							if (cellValor != null) {
								valorOriginal = cellValor.getNumericCellValue();
							}
							if (valorOriginal == 0) {
								valorProcessamento = 60;
							} else {
								valorProcessamento = valorOriginal;
							}

							Date dataFimPlaca = DbUtils.consultarPlaca(conexaoBD, placa);
							String resultado;
							String textoData;							
							if (dataFimPlaca != null) {
								if (dataAba.after(dataFimPlaca)) {
									resultado = "Inoperante";
									valorInoperante += valorProcessamento;
								} else {
									resultado = "Operante";
									valorOperante += valorProcessamento;
								}
								textoData = formatarData(dataFimPlaca);
							} else {
								resultado = "Não Verificável";
								// valorOperante += valor;
								valorNaoVerificavel += valorProcessamento;
								textoData = "Veículo não cadastrado.";
							}

							Cell cellResultado = linha.createCell(4);
							cellResultado.setCellValue(resultado);
							cellResultado.setCellStyle(estiloSemNegrito);

							Cell cellData = linha.createCell(5);
							cellData.setCellStyle(estiloSemNegrito);
							cellData.setCellType(CellType.STRING);
							cellData.setCellValue(textoData);

							System.out.println("Aba em processamento:" + formatarData(dataAba));
							System.out.println("Data fim placa encontrada:" + textoData);
							System.out.println("Placa:" + placa + " -  Resultado:" + resultado);

							DbUtils.inserirRegistroResultado(conexaoBD, dataAba, placa, null,
									getBigDecimal(valorOriginal), getBigDecimal(valorProcessamento),
									linha.getCell(2).getStringCellValue(), // buscando o Status original da planilha
									resultado, dataFimPlaca);
						}
					}
				}
				inserirStringNaPlanilha(abaAtual, getValorNumericoFormatado(valorOperante), coordenadaValorOperante,
						estiloSemNegrito);
				inserirStringNaPlanilha(abaAtual, getValorNumericoFormatado(valorInoperante), coordenadaValorInoperante,
						estiloSemNegrito);
				inserirStringNaPlanilha(abaAtual, getValorNumericoFormatado(valorNaoVerificavel),
						coordenadaValorNaoAceitoAscobom, estiloSemNegrito);

				double valorCobradoPagoAscobom = getValorNumericoPlanilha(abaAtual, "B4");
				double valorDevidoAscobom = getValorNumericoPlanilha(abaAtual, "E4");
				double valorTotalAscobom = getValorNumericoPlanilha(abaAtual, "B9");

				DbUtils.inserirTotal(conexaoBD, dataAba, getBigDecimal(valorOperante), getBigDecimal(valorInoperante),
						getBigDecimal(valorNaoVerificavel), getBigDecimal(valorCobradoPagoAscobom),
						getBigDecimal(valorDevidoAscobom), getBigDecimal(valorTotalAscobom));
				autoSize(abaAtual);
			}
		}
		FileOutputStream fileOut = new FileOutputStream(arquivo);
		wb.write(fileOut);
		wb.close();
		fileOut.close();
	}

	// Este método deve processar apenas o arquivo de 2011, e irá observar apenas as
	// abas a partir de Março.
	public static void processar2011Contrato(Connection conexaoBD, String arquivo, String aba) throws Exception {

		Workbook wb = new XSSFWorkbook(new FileInputStream(arquivo));
		Iterator<Sheet> sheetIterator = wb.sheetIterator();
		CellStyle estiloNegrito = getEstiloNegrito(wb);
		CellStyle estiloSemNegrito = getEstiloSemNegrito(wb);

		String coordenadaValorOperante = "F6";
		String coordenadaValorInoperante = "F7";
		String coordenadaValorNaoAceitoAscobom = "F8";

		while (sheetIterator.hasNext()) {
			Sheet abaAtual = sheetIterator.next();
			if (!abaAtual.getSheetName().equals("JANEIRO 2011") && !abaAtual.getSheetName().equals("FEVEREIRO 2011")
					&& abaAtual.getSheetName().equals(aba)) {

				java.util.Date dataAba = DbUtils.geraDataFimAba(abaAtual.getSheetName());
				// DateTimeFormatter formatter =
				// DateTimeFormatter.ofPattern("dd/MM/yyyy",Locale.US);
				Iterator<Row> linhaIterator = abaAtual.iterator();

				double valorOperante = 0;
				double valorInoperante = 0;
				double valorNaoVerificavel = 0;

				/*
				 * inserirStringNaPlanilha(abaAtual, "TOTAL OPERANTE", "E5");
				 * inserirStringNaPlanilha(abaAtual, "TOTAL INOPERANTE", "E6");
				 * inserirStringNaPlanilha(abaAtual, "NÃO ACEITO ASCOBOM", "E7");
				 */

				Cell valorTotalOperanteCell = null;
				Cell valorTotalInoperanteCell = null;
				Cell valorNaoAceitoAscobomCell = null;

				// Iterando nas linhas da aba aberta
				while (linhaIterator.hasNext()) {
					Row linha = linhaIterator.next();

					if (linha.getRowNum() == 6) {
						Cell labelTotalOperanteCell = linha.createCell(4);
						labelTotalOperanteCell.setCellStyle(estiloNegrito);
						labelTotalOperanteCell.setCellValue("TOTAL OPERANTE");
						valorTotalOperanteCell = linha.createCell(5);

					}
					if (linha.getRowNum() == 7) {
						Cell labelTotalInoperanteCell = linha.createCell(4);
						labelTotalInoperanteCell.setCellStyle(estiloNegrito);
						labelTotalInoperanteCell.setCellValue("TOTAL INOPERANTE");
						valorTotalInoperanteCell = linha.createCell(5);

					}
					if (linha.getRowNum() == 8) {
						Cell labelNaoAceitoAscobomCell = linha.createCell(4);
						labelNaoAceitoAscobomCell.setCellStyle(estiloNegrito);
						labelNaoAceitoAscobomCell.setCellValue("TOTAL NÃO VERIFICÁVEL");
						valorNaoAceitoAscobomCell = linha.createCell(5);
					}

					// Linha que cabeçalho deve ser criado
					if (linha.getRowNum() == 9) {
						criarCabecalhoResultadoPlanilha(wb, estiloNegrito, linha);
					}

					// Linha que inicia a ter dados para processamento
					if (linha.getRowNum() > 9) {
						System.out.println("Iniciando registro da aba:" + formatarData(dataAba));
						String contrato = "";
						String placa = "";

						// Buscando contrato
						Cell contratoCell = linha.getCell(0);
						if (contratoCell != null) {
							contratoCell.setCellType(CellType.STRING);
							contrato = contratoCell.getStringCellValue();
						}

						// Buscando placa
						Cell placaCell = linha.getCell(1);
						if (placaCell != null) {
							placaCell.setCellType(CellType.STRING);
							placa = placaCell.getStringCellValue();
						}

						if (!contrato.isEmpty() || !placa.isEmpty()) {
							// Buscando valor
							Cell cellValor = linha.getCell(3);
							double valorOriginal = 0.0;
							double valorProcessamento = 0.0;
							if (cellValor != null) {
								cellValor.setCellType(CellType.NUMERIC);
								valorOriginal = cellValor.getNumericCellValue();
							}
							if (valorOriginal == 0) {
								valorProcessamento = 60;
							} else {
								valorProcessamento = valorOriginal;
							}

							Date dataFimPlaca = null;
							Date dataFimContrato = null;
							Date dataPesquisa = null;

							if (!placa.isEmpty()) {
								dataFimPlaca = DbUtils.consultarPlaca(conexaoBD, placa);
							}
							if (!contrato.isEmpty()) {
								dataFimContrato = DbUtils.consultarContrato(conexaoBD, contrato);
							}
							if (dataFimPlaca != null && dataFimContrato != null) {
								System.out.println(
										"Data por contrato e placa foram encontradas. Escolhendo qual a menor;");
								if (dataFimPlaca.after(dataFimContrato)) {
									System.out.println("Utilizando data encontrada pelo contrato");
									dataPesquisa = dataFimContrato;
								} else {
									System.out.println("Utilizando data encontrada pela placa");
									dataPesquisa = dataFimPlaca;
								}
							} else {
								System.out
										.println("Apenas uma das datas está preenchida. Verificando qual utilizar...");
								if (dataFimPlaca != null) {
									System.out.println("Utilizando data encontrada pela placa");
									dataPesquisa = dataFimPlaca;
								}
								if (dataFimContrato != null) {
									System.out.println("Utilizando data encontrada pelo contrato");
									dataPesquisa = dataFimContrato;
								}
							}
							String resultado;
							String textoData;
						
							if (dataPesquisa != null) {
								if (dataAba.after(dataPesquisa)) {
									resultado = "Inoperante";
									valorInoperante += valorProcessamento;
								} else {
									resultado = "Operante";
									valorOperante += valorProcessamento;
								}
								textoData = formatarData(dataPesquisa);
							} else {
								resultado = "Não Verificável ";
								valorNaoVerificavel += valorProcessamento;
								textoData = "Veículo não cadastrado";								
							}
							String statusOriginal = "";
							Cell statusOriginalCell = linha.getCell(2);
							if (statusOriginalCell != null) {
								if (statusOriginalCell.getCellType().toString().equals("STRING")) {
									statusOriginal = statusOriginalCell.getStringCellValue();
								} else {
									statusOriginal = String.valueOf(statusOriginalCell.getNumericCellValue());
								}
							}
							System.out.println(
									"Placa:" + placa + " Contrato: " + contrato + " -  Resultado:" + resultado);
							System.out.println(
									"Total operante:" + valorOperante + " Total inoperante:" + valorInoperante);

							Cell cellResultado = linha.createCell(4);
							cellResultado.setCellValue(resultado);
							cellResultado.setCellStyle(estiloSemNegrito);

							Cell cellData = linha.createCell(5);
							cellData.setCellStyle(estiloSemNegrito);
							cellData.setCellValue(textoData);

							DbUtils.inserirRegistroResultado(conexaoBD, dataAba, placa, contrato,
									getBigDecimal(valorOriginal), getBigDecimal(valorProcessamento), statusOriginal,
									resultado, dataPesquisa);
						} else {
							System.out.println("Ambos placa e contrato estão vazios");
						}
						System.out.println("Processamento do registro finalizado");
					}
				}
				valorTotalOperanteCell.setCellValue(getValorNumericoFormatado(valorOperante));
				valorTotalInoperanteCell.setCellValue(getValorNumericoFormatado(valorInoperante));
				valorNaoAceitoAscobomCell.setCellValue(getValorNumericoFormatado(valorNaoVerificavel));
				valorTotalOperanteCell.setCellStyle(estiloNegrito);
				valorTotalInoperanteCell.setCellStyle(estiloNegrito);
				valorNaoAceitoAscobomCell.setCellStyle(estiloNegrito);

				double valorCobradoPagoAscobom = getValorNumericoPlanilha(abaAtual, "B4");
				double valorDevidoAscobom = getValorNumericoPlanilha(abaAtual, "D4");
				double valorTotalAscobom = getValorNumericoPlanilha(abaAtual, "B9");

				DbUtils.inserirTotal(conexaoBD, dataAba, getBigDecimal(valorOperante), getBigDecimal(valorInoperante),
						getBigDecimal(valorNaoVerificavel), getBigDecimal(valorCobradoPagoAscobom),
						getBigDecimal(valorDevidoAscobom), getBigDecimal(valorTotalAscobom));
				autoSize(abaAtual);
			}
		}
		FileOutputStream fileOut = new FileOutputStream(arquivo);
		wb.write(fileOut);
		wb.close();
		fileOut.close();
	}

	private static void criarCabecalhoResultadoPlanilha(Workbook wb, CellStyle estilo, Row linha) {

		Cell cellTitulo1 = linha.createCell(4);
		cellTitulo1.setCellValue("VALIDAÇÃO");
		cellTitulo1.setCellStyle(estilo);

		Cell cellTitulo2 = linha.createCell(5);
		cellTitulo2.setCellValue("DATA FIM");
		cellTitulo2.setCellStyle(estilo);
	}

	private static BigDecimal getBigDecimal(double valorOperante) {
		return new BigDecimal(valorOperante).setScale(2, RoundingMode.HALF_DOWN);
	}

	private static Double getValorNumericoPlanilha(Sheet abaAtual, String coordenada) {
		Cell cellTotalCobradoPagoAscobom = getReferenciaCelulaPlanilha(abaAtual, coordenada);
		cellTotalCobradoPagoAscobom.setCellType(CellType.NUMERIC);
		Double valorTotalCobrancaPagoAscobom = cellTotalCobradoPagoAscobom.getNumericCellValue();
		return valorTotalCobrancaPagoAscobom;
	}

	private static String getValorNumericoFormatado(double valorOperante) {
		NumberFormat nf = NumberFormat.getCurrencyInstance();
		return nf.format(getBigDecimal(valorOperante));
	}

	private static String formatarData(java.util.Date dataFimPlaca) {
		SimpleDateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
		return formato.format(dataFimPlaca);
	}

	private static Cell getReferenciaCelulaPlanilha(Sheet abaAtual, String coordenada) {
		CellReference cellReference = new CellReference(coordenada);
		Row row = abaAtual.getRow(cellReference.getRow());
		return row.getCell(cellReference.getCol());
	}

	private static void inserirStringNaPlanilha(Sheet abaAtual, String s, String coordenada, CellStyle style) {
		Cell cellPositivo = getReferenciaCelulaPlanilha(abaAtual, coordenada);
		cellPositivo.setCellValue(s);
		cellPositivo.setCellStyle(style);
	}

	private static CellStyle getEstiloNegrito(Workbook wb) {
		Font f1 = wb.createFont();
		CellStyle csNegrito = wb.createCellStyle();
		f1.setBold(true);
		f1.setFontHeightInPoints((short) 10);
		csNegrito.setFont(f1);
		return csNegrito;
	}

	private static CellStyle getEstiloSemNegrito(Workbook wb) {
		Font f1 = wb.createFont();
		CellStyle csNegrito = wb.createCellStyle();
		f1.setBold(true);
		f1.setFontHeightInPoints((short) 10);
		csNegrito.setFont(f1);
		return csNegrito;
	}

}