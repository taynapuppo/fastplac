
# ─────────────────────────────────────────────────────────────
#  IDs dos templates no Google Drive (mesmos do Apps Script)
# ─────────────────────────────────────────────────────────────
TEMPLATE_IDS = {
    "Placa Porta Paletes":        "1j9s7TPy_DhkSAWIFUoGRhnp4jtRdNu4Fpe3C4H7H26Q",
    "Placa PP Palete e Plano":    "1ypY-k2HkTWbfS-L2ePYTpD8-kdzR76i5QWkZTB-p1Rk",
    "Placa PP Peso Plano":        "14yuJSLFjlRdV6Ws6vPzVps4UPKEa7Ls2Ax72JmGu0bo",
    "Placa PP Peso por Nivel":    "1caWkpmv9XdYrpIWEOBeCwam5oi3FhD-xlg4ACBVd5sc",
    "Placa Mercado Livre":        "1rllpTeqLacZiqeCmd2X9sy1zNKgfTBVhHe2N0_TXi-o",
    "Placa Flow Rack":            "1QM2yzANFF1ltUCzfV_jA0QD1KTTkIeTm-HEKKQGKqjE",
    "Placa Drive In":             "137X648yX51Ebx83UhSOO7pCGdGYEoPSo4JQMoibY_fo",
    "Placa Mezanino":             "1rEbylWDSrhoygX02waFcLBt5kgLTnHWD_YCJNSc9KMo",
    "Placa Mezanino Carga Piso":  "1VgzTAC2n7yE3XuDF8tiP1lALWT6ISBA2aRr9p_yNFzM",
    "Placa Dinamico":             "1dJ7wkIA-gmdvOHCJ-wzjk-gCC-aztWlLIr0p17tCuk0",
    "Placa Dinamico Peso Maximo": "1ad88ne9EYxabZQfLW4_I8l-0biOx_cRs69nIXJA9jmU",
    "Placa Especial Furukawa":     "1zFYJPwazjsNBjkN12tR_ynegM4xyuobb2XLsKNgxl6g",
}

FOLDER_ID = "1S2mj8z1hv5sKmzXKaRcwUo20hdhIapDQ"

# ─────────────────────────────────────────────────────────────
#  Campos presentes em TODOS os tipos
# ─────────────────────────────────────────────────────────────
CAMPOS_COMUNS = [
    {"key": "Cliente",              "label": "Cliente",         "type": "text"},
    {"key": "N° do Projeto",        "label": "N° do Projeto",   "type": "text"},
    {"key": "N° do Pedido",         "label": "N° do Pedido",    "type": "text"},
    {"key": "Quantidade de Placas", "label": "Qtd. de Placas",  "type": "number"},
]

# ─────────────────────────────────────────────────────────────
#  Campos específicos por tipo  (ajuste conforme seus templates)
# ─────────────────────────────────────────────────────────────
CAMPOS_ESPECIFICOS = {
    "Placa Porta Paletes": [
        {"key": "Dimensões",                     "label": "Dimensões",                    "type": "text"},
        {"key": "Peso por Nível",                "label": "Peso por Nível",               "type": "text"},
        {"key": "N° de Níveis",                  "label": "N° de Níveis",                 "type": "text"},
        {"key": "Altura do 1° Nível",            "label": "Altura do 1° Nível",           "type": "text"},
    ],
    "Placa Flow Rack": [
        {"key": "Peso Máximo",                       "label": "Peso Máximo",                      "type": "text"},
        {"key": "N° de Níveis",                      "label": "N° de Níveis",                     "type": "text"},
         {"key": "N° de Caixas em Profundidade",     "label": "N° de Caixas em Profundidade",     "type": "text"},
        {"key": "Dimensões",                         "label": "Dimensões",                        "type": "text"},
    ],
    "Placa Drive In": [
        {"key": "Dimensões",                     "label": "Dimensões",                       "type": "text"},
        {"key": "Peso Máximo",                   "label": "Peso Máximo",                     "type": "text"},
        {"key": "N° de Níveis",                 "label": "N° de Níveis",                     "type": "text"},
        {"key": "N° de Paletes em Profundidade", "label": "N° de Paletes em Profundidade",   "type": "text"},
    ],
    "Placa Mezanino": [
        {"key": "Carga por m² (Piso)",       "label": "Carga por m² (Piso)",    "type": "text"},
        {"key": "Carga por Plano",           "label": "Carga por Plano",        "type": "text"},
        {"key": "N° de Níveis",              "label": "N° de Níveis",           "type": "text"},
        {"key": "N° de Pavimento",           "label": "N° de Pavimento",        "type": "text"},
    ],
    "Placa Mezanino Carga Piso": [
        {"key": "Carga por m² (Piso)",       "label": "Carga por m² (Piso)",    "type": "text"},
        {"key": "N° de Pavimento",           "label": "N° de Pavimento",        "type": "text"},
    ],
    "Placa Dinamico": [
        {"key": "Dimensões",                          "label": "Dimensões",                             "type": "text"},
        {"key": "Peso Máximo",                        "label": "Peso Máximo",                           "type": "text"},
        {"key": "Peso Mínimo",                        "label": "Peso Mínimo",                           "type": "text"},
        {"key": "N° de Níveis",                       "label": "N° de Níveis",                          "type": "text"},
        {"key": "N° de Paletes em Profundidade",      "label": "N° de Paletes em Profundidade",         "type": "text"},
    ],
    "Placa Mercado Livre": [
        {"key": "Dimensões",                   "label": "Dimensões",                  "type": "text"},
        {"key": "Peso Máximo",                 "label": "Peso Máximo",                "type": "text"},
        {"key": "Art de Projeto",              "label": "Art de Projeto",             "type": "text"},
        {"key": "Art de Montagem",             "label": "Art de Montagem",            "type": "text"},
    ],
    "Placa PP Palete e Plano": [
        {"key": "Dimensões",                     "label": "Dimensões",                    "type": "text"},
        {"key": "Peso por Plano",                "label": "Peso por Plano",               "type": "text"},
        {"key": "Peso por Palete",               "label": "Peso por Palete",              "type": "text"},
        {"key": "N° de Níveis",                  "label": "N° de Níveis",                 "type": "text"},
        {"key": "Altura do 1° Nível",            "label": "Altura do 1° Nível",           "type": "text"},
    ],
    "Placa PP Peso Plano": [
        {"key": "Peso por Plano",                "label": "Peso por Plano",               "type": "text"},
        {"key": "N° de Níveis",                  "label": "N° de Níveis",                 "type": "text"},
        {"key": "Altura do 1° Nível",            "label": "Altura do 1° Nível",           "type": "text"},
    ],
    "Placa PP Peso por Nivel": [
        {"key": "Dimensões",                     "label": "Dimensões",                    "type": "text"},
        {"key": "Carga Nível 1° ao 3°",          "label": "Carga Nível 1° ao 3°",         "type": "text"},
        {"key": "Carga Nível 4° ao 7°",          "label": "Carga Nível 4° ao 7°",         "type": "text"},
        {"key": "N° de Níveis",                  "label": "N° de Níveis",                 "type": "text"},
        {"key": "Altura do 1° Nível",            "label": "Altura do 1° Nível",           "type": "text"},
    ],
    "Placa Dinamico Peso Maximo": [
        {"key": "Dimensões",                              "label": "Dimensões",                              "type": "text"},
        {"key": "Peso Máximo",                            "label": "Peso Máximo",                            "type": "text"},
        {"key": "N° de Níveis",                           "label": "N° de Níveis",                           "type": "text"},
        {"key": "N° de Paletes em Profundidade",          "label": "N° de Paletes em Profundidade",          "type": "text"},
    ],
    "Placa Especial Furukawa": [
        {"key": "N° de Níveis",           "label": "N° de Níveis",                "type": "text"},
        {"key": "Altura do 1° Nível",     "label": "Altura do 1° Nível",          "type": "text"},
        {"key": "Nível 1",                "label": "Nível 1",                     "type": "text"},
        {"key": "Nível 2",                "label": "Nível 2",                     "type": "text"},
        {"key": "Nível 3",                "label": "Nível 3",                     "type": "text"},
        {"key": "Nível 4",                "label": "Nível 4",                     "type": "text"},
        {"key": "Nível 5",                "label": "Nível 5",                     "type": "text"},
        {"key": "Nível 6",                "label": "Nível 6",                     "type": "text"},       
    ]   
}
