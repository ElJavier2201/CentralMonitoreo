-- ============================================================
-- Central de Monitoreo para Simulaciones de Circuitos y Prototipos
-- Esquema de base de datos (dialecto Jet/ACE SQL de Microsoft Access)
--
-- Este archivo es de REFERENCIA y documentación. Para crear la base
-- de datos no es necesario ejecutarlo a mano: usa db/instalar.asp,
-- que aplica estas mismas sentencias mediante ADO sobre el archivo
-- CentralMonitoreo.accdb. Si prefieres crearla manualmente desde
-- Microsoft Access, pega cada bloque en la vista SQL de una consulta.
-- ============================================================

-- Tabla 1: cabecera del experimento
CREATE TABLE Proyectos (
    ID_Proyecto            COUNTER PRIMARY KEY,
    Nombre_Proyecto        VARCHAR(150) NOT NULL,
    Plataforma_Simulacion  VARCHAR(80)  NOT NULL,   -- Proteus, Tinkercad, etc.
    Microcontrolador       VARCHAR(80)  NOT NULL,   -- Cerebro del circuito (Arduino, ESP32...)
    Descripcion            VARCHAR(255),
    Fecha_Creacion         DATETIME
);

-- Tabla 2: inventario y mapa de cableado por proyecto
CREATE TABLE Componentes (
    ID_Componente          COUNTER PRIMARY KEY,
    ID_Proyecto            LONG NOT NULL,
    Tipo_Componente        VARCHAR(100) NOT NULL,   -- placa, resistencia, sensor, etc.
    Valor_Calculado        VARCHAR(100),            -- ej. "4.7 kOhm", "100 nF"
    Pin_Conexion           VARCHAR(50),             -- ej. "A0", "D2"
    Ubicacion_Protoboard   VARCHAR(100),            -- coordenadas exactas (fila/columna)
    Notas                  VARCHAR(255),
    CONSTRAINT FK_Componentes_Proyectos FOREIGN KEY (ID_Proyecto)
        REFERENCES Proyectos (ID_Proyecto)
);

-- Tabla 3: bitácora de fallos durante pruebas físicas
CREATE TABLE Bitacora_Fallos (
    ID_Fallo               COUNTER PRIMARY KEY,
    ID_Proyecto            LONG NOT NULL,
    Fecha_Registro         DATETIME,
    Sintoma_Error          VARCHAR(255) NOT NULL,   -- ej. lecturas inestables en monitor serie
    Solucion_Aplicada      VARCHAR(255),            -- solución electrónica o de código
    Estado                 VARCHAR(20),             -- "Pendiente" / "Resuelto"
    CONSTRAINT FK_Fallos_Proyectos FOREIGN KEY (ID_Proyecto)
        REFERENCES Proyectos (ID_Proyecto)
);
