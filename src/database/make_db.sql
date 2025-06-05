DROP TABLE export_line;
DROP TABLE export;

CREATE TABLE export (
    id AUTOINCREMENT PRIMARY KEY,
    status TEXT,
    created_at DATETIME
);

CREATE TABLE export_line (
    id AUTOINCREMENT PRIMARY KEY,
    Raison_sociale TEXT,
    Code_postal TEXT,
    Ville TEXT,
    Tiers_facture TEXT,
    Pays TEXT,
    Nature TEXT,
    Numero TEXT,
    Date_piece DATETIME,
    Date_Eche DATETIME,
    Devise TEXT,
    Total_TTC DOUBLE,
    Acompte DOUBLE,
    Net_a_payer DOUBLE,
    Conditions_de_reglement TEXT,
    Code_SIRET TEXT NULL,
    Secteur_activite TEXT,
    Commentaire TEXT NULL,
    Tiers TEXT,
    Modifiable TEXT,
    status TEXT,
    date_trait DATETIME NULL,
    date_reception_factor DATE NULL,
    date_compta DATE NULL,
    export_id INTEGER,
    CONSTRAINT FK_export_line_export FOREIGN KEY (export_id) REFERENCES export(id) ON DELETE CASCADE
);

