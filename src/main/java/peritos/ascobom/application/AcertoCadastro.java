package peritos.ascobom.application;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import peritos.ascobom.db.ConnectionFactory;

public class AcertoCadastro {   
	public static void main(String[] args) throws SQLException {
    Connection conexao = ConnectionFactory.getConnection();

    String sql = "select placa from cadastro_unificado "+
        "where placa is not null and placa <> '' "+
        "group by placa "+
        "having count(*)>1";
    PreparedStatement stmt = conexao.prepareStatement(sql);
    ResultSet rs = stmt.executeQuery();
    while (rs.next()) {
        String placa = rs.getString("placa");
        sql = "select contrato from cadastro_unificado where placa = ? and contrato is not null";
        PreparedStatement stmt2 = conexao.prepareStatement(sql);
        stmt2.setString(1, placa);
        ResultSet rsContrato = stmt2.executeQuery();

        if (rsContrato.next()) {
            String contrato =  rsContrato.getString("contrato");
            sql = "update cadastro_unificado " +
                "set contrato = ? " +
                "where placa = ? and contrato is NULL";
            PreparedStatement stmt3 = conexao.prepareStatement(sql);
            stmt3.setString(1,contrato);
            stmt3.setString(2,placa);
            System.out.println("Atualizada placa:"+placa+" inserindo o contrato:"+contrato);
            stmt3.execute();
            stmt3.close();
        }
        else {
            System.out.println("Falha grave. Contrato não encontrado para a "+placa);
        }
        stmt2.close();
    }
    stmt.close();
    conexao.close();
}
}