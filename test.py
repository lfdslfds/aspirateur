# Aspirateur
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

# Workbook is created
wb = xlsxwriter.Workbook('aspirateur.xlsx')

# add_sheet is used to create sheet.
sheet1 = wb.add_worksheet('Feuil1')

# Titres des colonnes
sheet1.write(0, 2, "Num_CMD")
sheet1.write(0, 3, "DPT")
sheet1.write(0, 4, "LIVRAISON")
sheet1.write(0, 5, "Adresse")
sheet1.write(0, 6, "CP")
sheet1.write(0, 7, "Ville")
sheet1.write(0, 8, "Tel1")
sheet1.write(0, 9, "Tel2")
sheet1.write(0, 10, "RESP")
sheet1.write(0, 11, "Tel")
sheet1.write(0, 12, "Email")
sheet1.write(0, 13, "Num_membre")
sheet1.write(0, 14, "Nom_membre")
sheet1.write(0, 15, "Email_membre")
sheet1.write(0, 16, "Trie")
sheet1.write(0, 17, "Code")
sheet1.write(0, 18, "Designation")
sheet1.write(0, 19, "Prix unitaire")
sheet1.write(0, 20, "Qte")
sheet1.write(0, 21, "Prix")
sheet1.write(0, 22, "FDP")
sheet1.write(0, 23, "Prix sans FDP")

# definition variable
catalogue = int
membre = int
produit = int

# Commancer à écrire sur la ligne 2
# catalogue = 1
last_row_data = 1

# Boucle catalogue
for catalogue in range(1, 61):
    catalogue_str = str(catalogue)

    # Boucle membre
    for membre in range(49, 129, 4):
        # Definition infos membre
        membre_num = xl_rowcol_to_cell(4, membre - 1)
        membre_num_str = str(membre_num)
        membre_nom = xl_rowcol_to_cell(9, membre - 1)
        membre_nom_str = str(membre_nom)
        membre_email = xl_rowcol_to_cell(10, membre - 1)
        membre_email_str = str(membre_email)

        # Boucle produit
        for produit in range(1, 301):
            # Defintion infos produits
            produit = produit + 11
            produit_str = str(produit)
            trie = xl_rowcol_to_cell(produit, 20)
            trie_str = str(trie)
            produit_qty = xl_rowcol_to_cell(produit - 1, membre - 1)
            produit_qty_str = str(produit_qty)
            produit_prix = xl_rowcol_to_cell(produit - 1, membre + 2)
            produit_prix_str = str(produit_prix)
            last_row_data_str = str(last_row_data + 1 )

            sheet1.write(last_row_data, 2,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str +'.xlsx]ADMIN1\'!$AI$4')
            sheet1.write(last_row_data, 3,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!DEPT')
            sheet1.write(last_row_data, 4,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$X$3')
            sheet1.write(last_row_data, 5,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$X$4')
            sheet1.write(last_row_data, 6,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$X$5')
            sheet1.write(last_row_data, 7,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$X$6')
            sheet1.write(last_row_data, 8,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$X$7')
            sheet1.write(last_row_data, 9,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$AD$7')
            sheet1.write(last_row_data, 10,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$X$8')
            sheet1.write(last_row_data, 11,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$X$9')
            sheet1.write(last_row_data, 12,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$X$10')
            sheet1.write(last_row_data, 13,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!' + membre_num_str)
            sheet1.write(last_row_data, 14,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!' + membre_nom_str)
            sheet1.write(last_row_data, 15,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!' + membre_email_str)
            sheet1.write(last_row_data, 16,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!' + trie_str)
            sheet1.write(last_row_data, 17,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$V$' + produit_str)
            sheet1.write(last_row_data, 18,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$W$' + produit_str)
            sheet1.write(last_row_data, 19,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$AK$' + produit_str)
            sheet1.write(last_row_data, 20,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!' + produit_qty_str)
            sheet1.write(last_row_data, 21,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!' + produit_prix_str)
            sheet1.write(last_row_data, 22,
                         '=\'D:\\FRUITSTOCK\\KDrive\\Common documents\\04_SUIVI_COMMANDE\\07_ASPIRATEUR\\Televersement_cmd\\[cmd_' + catalogue_str + '.xlsx]ADMIN1\'!$AL$7')
            sheet1.write(last_row_data, 23,'=$V$' + last_row_data_str + '-($W$' + last_row_data_str + '*$U$' + last_row_data_str + ')')


            last_row_data = last_row_data + 1


wb.close()
