package peritos.ascobom.db;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class ConnectionFactory {
    public static Connection getConnection() throws SQLException{
        String url = "jdbc:mysql://localhost:3306/ascobom?useTimezone=true&serverTimezone=UTC"; //Nome da base de dados
        //String url = "jdbc:mysql://localhost:3306/ascobom";

        String user = "root"; //nome do usuário do MySQL trabalho (no notebook é hudson)
        String password = "123456"; //senha do MySQL notebook
         
        Connection conexao = DriverManager.getConnection(url, user, password);
         
        return conexao;
    }

}
