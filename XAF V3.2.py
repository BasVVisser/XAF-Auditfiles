#XAF V3.2 XML inlezen en exporteren naar CSV
#Dit programma leest  XAF V3.2 bestand in een Pandas Dataframe en maakt van alle tabellen 1 totaaltabel en exporteert deze naar een CSV-bestand.

#Benodigde modules
import pandas as pd
import numpy as np
import sys
import os
import xml.etree.ElementTree as ET

#Selecteren van bestanden, als dit via IDEA moet kan dit weg.
import tkinter as tk
from tkinter import filedialog

#Algemene functies
def XAF_Parsen(file):
    tree = ET.parse(file)
    root = tree.getroot()
    return root

def XAF_element_vinden(root, tag, ns):
    return root.find(tag, ns)

def namespace_ombouwen_transacties(root):
    ns = root.tag.split('{')[1].split('}')[0]
    ns = "{" + ns + "}"
    return ns

def namespace_ombouwen_algemeen(root):
    namespaces = {'xsd':"http://www.w3.org/2001/XMLSchema", 'xsi':"http://www.w3.org/2001/XMLSchema-instance" }
    namespaces["ADF"] = root.tag.split('{')[1].split('}')[0]
    return namespaces

def hoofdlaag_informatie(root, ns):
    if root == None:
        return None
    else: 
        informatie_opslaan = dict()
        for child in root:
            kolomnaam = child.tag.replace(ns,'')
            kolomwaarde = child.text
            if len(child) == 0:
                informatie_opslaan[kolomnaam] = kolomwaarde     
            else:
                continue
        return informatie_opslaan

def informatie_tweede_laag(root, ns):
    huidige_regel = 0
    regels = dict()

    for child in root:
        informatie_opslaan = dict()
        for subchild in child:
            if len(subchild) != 0:
                for subsubchild in subchild:
                    kolomwaarde = subsubchild.text
                    kolomnaam = subsubchild.tag.replace(ns,'')
                    informatie_opslaan[kolomnaam] = kolomwaarde
            else:
                kolomwaarde =subchild.text
                kolomnaam = subchild.tag.replace(ns,'')
                informatie_opslaan[kolomnaam] = kolomwaarde
        regels[huidige_regel] = informatie_opslaan
        huidige_regel +=1
    df = pd.DataFrame(regels).transpose()
    return df

def btw_codes_oplossen(vatcode):
    claim = vatcode[(['vatID', 'vatDesc','vatToClaimAccID'])]
    claim = claim[pd.isnull(claim['vatToClaimAccID']) == False]
    pay = vatcode[(['vatID', 'vatDesc','vatToPayAccID'])]
    pay = pay[pd.isnull(pay['vatToPayAccID']) == False]
    vatcode = pd.merge(claim,pay, on = ['vatID', 'vatDesc'], how ='outer')
    return vatcode

def dagboek_informatie(dagboeken, ns):
    dagboek_df = pd.DataFrame()
    for dagboek in dagboeken:
        dagboekinfo = dict()
        for regels in dagboek:
            if len(regels) == 0:
                kolomnaams = regels.tag.replace(ns,'')
                kolomwaardes = regels.text
                dagboekinfo[kolomnaams] = kolomwaardes
        dagboek_df = dagboek_df.append( dagboekinfo, ignore_index = True)
    return dagboek_df

def transactie_informatie(dagboeken, ns):
    transacties_df = pd.DataFrame()
    total_records = list()
    record_dict = dict()
    for dagboek in dagboeken: 
        for regels in dagboek:
            if len(regels) == 0:
                kolomnaams = regels.tag.replace(ns,'')
                kolomwaardes = regels.text
                record_dict[kolomnaams] = kolomwaardes
            else:
                for record in regels: 
                    if len(record) == 0:
                        kolomnaams = "TR_"+record.tag.replace(ns,'')
                        kolomwaardes = record.text
                        record_dict[kolomnaams] = kolomwaardes 
                    else:
                        for subfields in record: 
                            if len(subfields) == 0:
                                kolomnaams = subfields.tag.replace(ns,'')
                                kolomwaardes = subfields.text
                                record_dict[kolomnaams] = kolomwaardes 
                            else: 

                                for subfields_1 in subfields: 
                                    if len(subfields_1) == 0:
                                        kolomnaams = subfields_1.tag.replace(ns,'')
                                        kolomwaardes = subfields_1.text
                                        record_dict[kolomnaams] = kolomwaardes
                                    else : print('nog een sublaag!')
                        total_records.append(record_dict.copy()) 
    transacties_df = transacties_df.append(total_records, ignore_index = True)
    return transacties_df

def openingsbalans_samenvoegen(transacties_df, openingsbalans, openingbalance):
    if openingsbalans == None:
        return transacties_df
    else:
        frames = pd.concat([openingbalance, transacties_df], axis = 0)
        return frames

def bedrag_corrigeren(transacties_df):
    amount_raw = transacties_df['amnt'].astype(float)
    conditions = [
        (transacties_df['amntTp'] == 'C'),
        (transacties_df['amntTp'] == 'D')]
    choices = [-1,1]
    transacties_df['amount'] = np.select(conditions, choices, default= 1 ) * amount_raw
    return transacties_df

def debet_toevoegen(transacties_df):
    amount_raw = transacties_df['amnt'].astype(float)
    conditions = [
        (transacties_df['amntTp'] == 'C'),
        (transacties_df['amntTp'] == 'D')]
    choices = [0,1] #credit is 0 1e keuze, debet is 1.
    transacties_df['debet'] = np.select(conditions, choices, default= 1 ) * amount_raw
    return transacties_df

def credit_toevoegen(transacties_df):
    amount_raw = transacties_df['amnt'].astype(float)
    conditions = [
        (transacties_df['amntTp'] == 'C'),
        (transacties_df['amntTp'] == 'D')]
    choices = [1,0]
    transacties_df['credit'] = np.select(conditions, choices, default= 1 ) * amount_raw
    return transacties_df

def btw_bedrag_corrigeren(transacties_df):
    if "vatAmnt" in transacties_df.columns:
        btw_bedrag_raw = transacties_df["vatAmnt"].astype(float)
        conditions = [
            (transacties_df["vatAmntTp"] == "C"),
            (transacties_df["vatAmntTp"] == "D")]
        choices = [-1,1]
        transacties_df["vat_amount"] = np.select(conditions, choices, default=1) * btw_bedrag_raw
        return transacties_df
    else:
        return transacties_df

def algemene_informatie_samenvoegen(openingsbalans, openingsbalansinfo, headerinfo, companyinfo, transactionsinfo):
    if openingsbalans == None:
        auditfile_algemene_informatie = pd.concat([headerinfo, companyinfo, transactionsinfo], axis = 1 )
        return auditfile_algemene_informatie
    else:
        auditfile_algemene_informatie = pd.concat([headerinfo, companyinfo, transactionsinfo, openingsbalansinfo], axis = 1)
        return auditfile_algemene_informatie

def tabellen_samenvoegen(transacties_df, periods, vatcode, custsup, genledg, dagboek_df):
    if periods.empty:
        transacties_df = transacties_df
    else:
        transacties_df = pd.merge(transacties_df, periods, left_on="TR_periodNumber",right_on = "periodNumber", how="left")
    
    if len(vatcode) != 0 and "vatID" in transacties_df.columns:
        transacties_df = pd.merge(transacties_df, vatcode, on="vatID", how="left")
    else:
        transacties_df = transacties_df

    if "custSupID" in transacties_df.columns and len(custsup) != 0:
        transacties_df = pd.merge(transacties_df, custsup.add_prefix("cs_"), left_on="custSupID", right_on="cs_custSupID", how="left")
    else:
        transacties_df = transacties_df

    if "accID" in transacties_df.columns:
        transacties_df = pd.merge(transacties_df, genledg, on="accID", how="left")
    elif "accountID" in transacties_df.columns:
        transacties_df = pd.merge(transacties_df, genledg, on="accountID", how="left")

    if "jrnID" in transacties_df.columns:
        transacties_df = pd.merge(transacties_df, dagboek_df.add_prefix("jrn_"), left_on="jrnID", right_on="jrn_jrnID", how="left")
    elif "journalID" in transacties_df.columns:
        transacties_df = pd.merge(transacties_df, dagboek_df.add_prefix("jrn_"), left_on="journalID", right_on="jrn_journalID", how="left")

    return transacties_df

def exportlocatie_bepalen(file):
    exportlocatie = file
    exportlocatie = exportlocatie[::-1].split("/",1)
    exportpad = exportlocatie[1][::-1]
    exportnaam = exportlocatie[0][::-1].strip(".xaf")
    exportbestand = str(exportpad)+"/"+str(exportnaam)+".csv"
    return exportbestand

def exportlocatie__geconsolideerd_bepalen(file):
    exportlocatie = file
    exportlocatie = exportlocatie[::-1].split("/",1)
    exportpad = exportlocatie[1][::-1]
    #exportnaam = exportlocatie[0][::-1].strip(".xaf")
    exportbestand = str(exportpad)+"/"+"Geconsolideerd.csv"
    return exportbestand

def dataframe_opschonen(transacties_df):
    if "amnt" in transacties_df:
        transacties_df = transacties_df.drop(columns=['amnt'])
    transacties_df = transacties_df.dropna(axis=1, how="all")
    return transacties_df

def openingsbalans_gegevens_toevoegen(openingbalance):
    if "jrnID" not in openingbalance.columns:
        openingbalance["jrnID"] = "BB"
    if "desc" not in openingbalance.columns:
        openingbalance["desc"] = "Beginbalans"
    if "TR_desc" not in openingbalance.columns:
        openingbalance["TR_desc"] = "Beginbalans"
    if "TR_periodNumber" not in openingbalance.columns:
        openingbalance["TR_periodNumber"] = 0
    #if "jrn_desc" not in openingbalance.columns:
    #    openingbalance["jrn_desc"] = "Beginbalans"
    return openingbalance

def openingsbalans_toevoegen_dagboeken(dagboek_df):
    openingsbalans_gegevens = {"jrnID":["BB"], "desc":["Beginbalans"]}
    tijdelijk = pd.DataFrame(data=openingsbalans_gegevens)
    dagboek_df = dagboek_df.append(tijdelijk)
    return dagboek_df

#Start hoofdprogramma
if __name__ == "__main__":
    
    main = tk.Tk()
    filez = filedialog.askopenfilenames(filetypes=[("XAF","*.xaf")],multiple=True)
    filenames = main.tk.splitlist(filez) 
    main.destroy()

    geconsolideerd = pd.DataFrame()

    for filename in filenames:
        file = filename

        #Parsen XML bestand
        root = XAF_Parsen(file)

        if "3.2" in root.tag:
            namespaces = namespace_ombouwen_algemeen(root)
            ns = namespace_ombouwen_transacties(root)
            
            #Algemene hoofdinformatie auditfile verkrijgen & samenvoegen
            header = XAF_element_vinden(root, "ADF:header", namespaces)
            headerinfo = pd.DataFrame(hoofdlaag_informatie(header, ns), index = [0]) 

            company = XAF_element_vinden(root, "ADF:company", namespaces)
            companyinfo = pd.DataFrame(hoofdlaag_informatie(company, ns), index = [0])

            transactions = XAF_element_vinden(root, "ADF:company/ADF:transactions", namespaces)
            transactionsinfo = pd.DataFrame(hoofdlaag_informatie(transactions, ns), index = [0])

            openingsbalans = XAF_element_vinden(root, "ADF:company/ADF:openingBalance", namespaces)
            openingsbalansinfo = pd.DataFrame(hoofdlaag_informatie(openingsbalans, ns), index = [0])

            auditfile_algemene_informatie = algemene_informatie_samenvoegen(openingsbalans, openingsbalansinfo, headerinfo, companyinfo, transactionsinfo)
        
            #Tabellen tweede laag informatie krijgen (algemene tabellen zoals klanten/leveranciers/grootboeknummers)
            openingbalance = informatie_tweede_laag(company.findall("ADF:openingBalance/ADF:obLine",namespaces),ns)
            periods = informatie_tweede_laag(company.findall("ADF:periods/ADF:period",namespaces),ns)
            custsup = informatie_tweede_laag(company.findall("ADF:customersSuppliers/ADF:customerSupplier",namespaces),ns)
            vatcode = informatie_tweede_laag(company.findall("ADF:vatCodes/ADF:vatCode",namespaces),ns)
            genledg = informatie_tweede_laag(company.findall("ADF:generalLedger/ADF:ledgerAccount",namespaces),ns)
            basics = informatie_tweede_laag(company.findall("ADF:generalLedger/ADF:basics",namespaces),ns)
            
            #Dagboekinformatie en transacties krijgen
            dagboeken = company.findall("ADF:transactions/ADF:journal", namespaces)
            dagboek_df = dagboek_informatie(dagboeken, ns)
            transacties_df = transactie_informatie(dagboeken, ns)

            #Tabellen tweedelaagsinformatie aanpassen om samen te voegen.
            vatcode = btw_codes_oplossen(vatcode)
            openingbalance = openingsbalans_gegevens_toevoegen(openingbalance)
            dagboek_df = openingsbalans_toevoegen_dagboeken(dagboek_df)

            #Samenvoegen van de transacties met de algemene tabellen + openingsbalans
            transacties_df = openingsbalans_samenvoegen(transacties_df,openingsbalans,openingbalance)
            transacties_df = bedrag_corrigeren(transacties_df)
            transacties_df = btw_bedrag_corrigeren(transacties_df)

            transacties_df = tabellen_samenvoegen(transacties_df, periods, vatcode, custsup, genledg, dagboek_df)

            #Debet/Credit kolom toevoegen, amnt verwijderen (is al gecorrigeerd naar amount), 
            transacties_df = debet_toevoegen(transacties_df)
            transacties_df = credit_toevoegen(transacties_df)
            transacties_df = dataframe_opschonen(transacties_df)

            exportbestand = exportlocatie_bepalen(file)
            
            transacties_df.to_csv(exportbestand,
                                index=False,
                                decimal=",",
                                sep=";")

            ##### Voor het geconsolideerde gedeelte, alles hiervoor was enkelvoudig.
            #Als er meer dan 1 bestand geselecteerd is wordt de data aan de geconsolideerde versie toegevoegd.
            if len(filenames) > 1:
                geconsolideerd = pd.concat([geconsolideerd, transacties_df], axis=0, ignore_index=True)



    #exporteren van de geconsolideerde ADF
    if len(filenames) > 1:
        exportbestand_geconsolideerd = exportlocatie__geconsolideerd_bepalen(filenames[0])
        geconsolideerd.to_csv(exportbestand_geconsolideerd,
                            index=False,
                            decimal=",",
                            sep=";")

