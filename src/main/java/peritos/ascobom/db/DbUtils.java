package peritos.ascobom.db;

import java.math.BigDecimal;
import java.sql.Connection;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Types;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Calendar;

public class DbUtils {
	public static Date consultarPlaca(Connection connection, String placa) throws SQLException {
		String sql = "select min(data_fim) data from cadastro_unificado where placa = ?";
		if (placa == null) {
			return null;
		}
		if (placa.length() < 7) {
			return null;
		}
		PreparedStatement stmt = connection.prepareStatement(sql);
		stmt.setString(1, placa);
		Calendar calendar = Calendar.getInstance();
		ResultSet rs = stmt.executeQuery();
		if (rs.next() && (rs.getDate("data") != null)) {
			System.out.println("Placa:" + placa + " data encontrada na base:" + rs.getString("data"));
			Date resultado = rs.getDate("data", calendar);
			rs.close();
			stmt.close();
			return resultado;
		} else {
			System.out.println("Placa não encontrada:" + placa);
			return null;
		}
	}

	public static Date consultarContrato(Connection connection, String contrato) throws SQLException {
		String sql = "select min(data_fim) data from cadastro_unificado where contrato = ?";
		PreparedStatement stmt = connection.prepareStatement(sql);
		stmt.setString(1, contrato);
		Calendar calendar = Calendar.getInstance();
		ResultSet rs = stmt.executeQuery();
		if (rs.next() && (rs.getDate("data") != null)) {
			System.out.println("Contrato:" + contrato + " data encontrada na base:" + rs.getString("data"));
			Date resultado = rs.getDate("data", calendar);
			rs.close();
			stmt.close();
			return resultado;
		} else {
			System.out.println("Contrato não encontrado:" + contrato);
			return null;
		}
	}

	public static LocalDate geraDataFimAbaLocalDate(String texto) throws ParseException {
		String nomeAba = getPrimeiraPalavra(texto);
		String anoAba = getUltimaPalavra(texto);
		int mes;
		switch (nomeAba) {
		case "JANEIRO":
			mes = 1;
			break;
		case "FEVEREIRO":
			mes = 2;
			break;
		case "MARÇO":
			mes = 3;
			break;
		case "ABRIL":
			mes = 4;
			break;
		case "MAIO":
			mes = 5;
			break;
		case "JUNHO":
			mes = 6;
			break;
		case "JULHO":
			mes = 7;
			break;
		case "AGOSTO":
			mes = 8;
			break;
		case "SETEMBRO":
			mes = 9;
			break;
		case "OUTUBRO":
			mes = 10;
			break;
		case "NOVEMBRO":
			mes = 11;
			break;
		case "DEZEMBRO":
			mes = 12;
			break;
		default:
			throw new IllegalStateException("Valor inesperado " + nomeAba);
		}
		// SimpleDateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
		// return formato.parse("01/" + mes + "/" + anoAba);
		return LocalDate.of(Integer.parseInt(anoAba), mes, 01);
	}

	public static java.util.Date geraDataFimAba(String texto) throws ParseException {
		String nomeAba = getPrimeiraPalavra(texto);
		String anoAba = getUltimaPalavra(texto);
		int mes;
		switch (nomeAba) {
		case "JANEIRO":
			mes = 1;
			break;
		case "FEVEREIRO":
			mes = 2;
			break;
		case "MARÇO":
			mes = 3;
			break;
		case "ABRIL":
			mes = 4;
			break;
		case "MAIO":
			mes = 5;
			break;
		case "JUNHO":
			mes = 6;
			break;
		case "JULHO":
			mes = 7;
			break;
		case "AGOSTO":
			mes = 8;
			break;
		case "SETEMBRO":
			mes = 9;
			break;
		case "OUTUBRO":
			mes = 10;
			break;
		case "NOVEMBRO":
			mes = 11;
			break;
		case "DEZEMBRO":
			mes = 12;
			break;
		default:
			throw new IllegalStateException("Valor inesperado " + nomeAba);
		}
		// return LocalDate.of(Integer.parseInt(anoAba),mes,1);
		SimpleDateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
		return formato.parse("01/" + mes + "/" + anoAba);
	}

	public static String getPrimeiraPalavra(String texto) {
		String[] s = texto.trim().split(" ");
		return s[0];
	}

	public static String getUltimaPalavra(String texto) {
		String[] s = texto.trim().split(" ");
		return s[1];
	}

	public static void inserirTotal(Connection connection, java.util.Date dataAba, BigDecimal valorOperante,
			BigDecimal valorInoperante, BigDecimal valorNaoReconhecidoAscobom, BigDecimal valorCobradoPagoAscobom,
			BigDecimal valorDevidoAscobom, BigDecimal valorTotalAscobom) throws Exception {
		String sql = "insert into total_mensal (data_aba,valor_operante,valor_inoperante,valor_nao_verificavel,valor_cobrado_pago_ascobom,valor_devido_ascobom,valor_total_ascobom) values (?,?,?,?,?,?,?)";
		PreparedStatement stmt = connection.prepareStatement(sql);
		stmt.setDate(1, new Date(dataAba.getTime()));
		stmt.setBigDecimal(2, valorOperante);
		stmt.setBigDecimal(3, valorInoperante);
		stmt.setBigDecimal(4, valorNaoReconhecidoAscobom);
		stmt.setBigDecimal(5, valorCobradoPagoAscobom);
		stmt.setBigDecimal(6, valorDevidoAscobom);
		stmt.setBigDecimal(7, valorTotalAscobom);
		stmt.execute();
		stmt.close();
	}

	public static void inserirRegistroResultado(Connection conexao, java.util.Date dataAba, String placa,
			String contrato, BigDecimal valorOriginal, BigDecimal valorProcessamento, String statusOriginal,
			String resultado, java.util.Date dataFimResultado) throws Exception {
		String sql = "insert into lancamento_mensal (data_aba,placa,contrato,status_original,resultado_processamento,data_fim_resultado,valor_original,valor_processamento) values (?,?,?,?,?,?,?,?)";
		PreparedStatement stmt = conexao.prepareStatement(sql);
		if (dataAba != null) {
			stmt.setDate(1, new Date(dataAba.getTime()));
		} else {
			stmt.setNull(1, Types.DATE);
		}
		stmt.setString(2, placa);
		stmt.setString(3, contrato);

		stmt.setString(4, statusOriginal);
		stmt.setString(5, resultado);
		if (dataFimResultado != null) {
			stmt.setDate(6, new Date(dataFimResultado.getTime()));
		} else {
			stmt.setNull(6, Types.DATE);
		}		
		stmt.setBigDecimal(7, valorOriginal);
		stmt.setBigDecimal(8, valorProcessamento);
		stmt.execute();
		stmt.close();
	}

}
