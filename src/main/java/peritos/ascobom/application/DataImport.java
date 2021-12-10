package peritos.ascobom.application;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import peritos.ascobom.db.ConnectionFactory;

public class DataImport {
	public static void ProcessarArquivoInoperante(){
        try {
            //Abrir uma determinada planilha e fazer a varredura.
            //Para cada linha encontrada, fazer a verificação(se insert ou update)

            //Primeiro: Ler planilha Equipamento inoperante.

            //Instanciando um objeto ligado ao arquivo xlsx, conforme desejado
            Connection conexao = ConnectionFactory.getConnection();

            String arquivo = "C:\\Dev\\Arquivos\\inoperante.xlsx";
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(arquivo));

            String placa;
            Date data; //Último report

            //Buscando a primeira planilhado do arquivo,índice zero
            Sheet sheetAlunos = wb.getSheetAt(0);

            Iterator<Row> rowIterator = sheetAlunos.iterator();
            int contador = 0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //Iterator<Cell> cellIterator = row.cellIterator();

                if (row.getRowNum() > 0) {
                    Cell celPlaca = row.getCell(0);
                    Cell celData = row.getCell(3);

                    celData.setCellType(CellType.NUMERIC);
                    celPlaca.setCellType(CellType.STRING);
                    //System.out.println("Linha:"+row.getRowNum()+" - Coluna:"+cell2.getColumnIndex()+" - Valor Célula:"+cell2.getStringCellValue());

                    data = celData.getDateCellValue();
                    placa = celPlaca.getStringCellValue();

                    DateFormat formatador = new SimpleDateFormat("dd/MM/yyyy");
                    System.out.println("Placa: " + placa + " Data: " +formatador.format(data));

                    insereRegistro(conexao,placa,null,data,1);
                }
                contador++;
            }
            System.out.println("Total de linhas processadas:"+contador);
            wb.close();
            conexao.close();

        }
        catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("Arquivo Excel não encontrado!");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    public static void ProcessarArquivoTermo(){
        try {

            Connection conexao = ConnectionFactory.getConnection();

            //Instanciando um objeto ligado ao arquivo xlsx, conforme desejado
            String arquivo = "C:\\Dev\\Arquivos\\termoCancelamento.xlsx";
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(arquivo));
            String contrato;
            String placa;
            Date data;

            //Buscando a primeira planilhado do arquivo,índice zero
            Sheet sheetAlunos = wb.getSheetAt(0);

            Iterator<Row> rowIterator = sheetAlunos.iterator();
            int contador = 0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //Iterator<Cell> cellIterator = row.cellIterator();

                if (row.getRowNum()>0) {
                    Cell celContrato = row.getCell(1);
                    Cell celPlaca= row.getCell(2);

                    celContrato.setCellType(CellType.STRING);
                    celPlaca.setCellType(CellType.STRING);
                    //System.out.println("Linha:"+row.getRowNum()+" - Coluna:"+cell2.getColumnIndex()+" - Valor Célula:"+cell2.getStringCellValue());

                    contrato = celContrato.getStringCellValue(); //Ver na aula  Algaworks como trabalhar com datas;
                    placa = celPlaca.getStringCellValue();
                    //Validar aqui se a placa é válida. Deve ter 3 letras e 4 números.
                    System.out.println("Placa: "+placa+" Contrato: "+contrato);

                    SimpleDateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
                    Date dataTermo = formato.parse("19/08/2011");
                    insereRegistro(conexao,placa,contrato,dataTermo,2);
                }
                contador++;
            }
            System.out.println("Total de linhas processadas: "+ contador);
            // FileOutputStream arquivoSaida = new FileOutputStream("/home/erebor/Desktop/ASCOBOM/termoCancelamentoImport.xlsx");
            // wb.write(arquivoSaida);
            wb.close();
            conexao.close();

        }
        catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("Arquivo Excel não encontrado!");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        } catch (ParseException e) {
            e.printStackTrace();
        }
    }

    public static void insereRegistro(Connection connection, String placa, String contrato,Date data, int origem) throws SQLException {
    
        String sql = "insert into cadastro_unificado (placa,contrato,data_fim,origem) values (?,?,?,?)";
        PreparedStatement stmt = connection.prepareStatement(sql);
        stmt.setString(1, placa);
        stmt.setString(2, contrato);
        stmt.setDate(3,new java.sql.Date(data.getTime()));
        stmt.setInt(4,origem);
        stmt.execute();
        stmt.close();
    }
    
    public static void main(String[] args) {
        ProcessarArquivoInoperante();
        ProcessarArquivoTermo();
    }
}
