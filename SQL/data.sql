-- ascobom.cadastro_unificado definition

CREATE TABLE `cadastro_unificado` (
  `placa` varchar(30) DEFAULT NULL,
  `contrato` varchar(30) DEFAULT NULL,
  `data_fim` date DEFAULT NULL,
  `origem` int DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;