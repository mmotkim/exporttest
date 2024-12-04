import * as PDFLib from "pdf-lib";
import fs from "fs";

const requiredHeadersData = [
  { column: "Matnr", data: null, columnName: "Material" },
  { column: "Maktx", data: null, columnName: "Material Description" },
  { column: "Matkl", data: null, columnName: "Material Group" },
  { column: "Mtart", data: null, columnName: "Material Type" },
  { column: "Bwkey", data: null, columnName: "Plant" },
  { column: "Spart", data: null, columnName: "Division" },
  { column: "Labst_E", data: null, columnName: "Base Unit of Measure" },
  { column: "Brgew", data: null, columnName: "Gross Weight" },
  { column: "Ntgew", data: null, columnName: "Net Weight" },
  { column: "Bdmng_E", data: null, columnName: "Weight Unit" },
];

const requiredMeasuresData = [
  { column: "Lgort", data: null, columnName: "Storage Location" },
  { column: "Labst", data: null, columnName: "Unrestricted Stock" },
  { column: "Bdmng", data: null, columnName: "Planned Quantity" },
  { column: "Avail", data: null, columnName: "Available Stock" },
];

const data = {
  columns: [
    {
      header: "Material",
      key: "Matnr",
      width: 15,
    },
    {
      header: "Material Group",
      key: "Maktx",
      width: 15,
    },
    {
      header: "Material Type",
      key: "Mtart",
      width: 15,
    },
    {
      header: "Company Code",
      key: "Bukrs",
      width: 15,
    },
    {
      header: "Plant",
      key: "Bwkey",
      width: 15,
    },
    {
      header: "Storage Location",
      key: "Lgort",
      width: 15,
    },
    {
      header: "Quantity in SLOC",
      key: "Labst",
      width: 15,
    },
    {
      header: "Reserved Stock",
      key: "Bdmng",
      width: 15,
    },
  ],
  rows: [
    ["DRYR1000", "Yuki Dryer 1000", "YKTG", "YKVN", "PLHN", "SLDG", "0", "0"],
    ["DRYR1000", "Yuki Dryer 1000", "YKTG", "YKVN", "PLSG", "SLDG", "0", "0"],
    ["FRZR1000", "Yuki Freezer", "YKTG", "YKVN", "PLHN", "SLDG", "0", "0"],
    ["RFGT1000", "Yuki Refrigerator", "YKTG", "YKVN", "PLHN", "SLDG", "0", "0"],
    ["WSMC1000", "Yuki Washing Machine", "HAWA", "YKVN", "PLHN", "SLDG", "0", "0"],
    ["DRYR1000", "Yuki Dryer 1000", "YKTG", "YKVN", "PLHN", "SLDO", "0", "0"],
    ["DRYR1000", "Yuki Dryer 1000", "YKTG", "YKVN", "PLSG", "SLDO", "0", "0"],
    ["FRZR1000", "Yuki Freezer", "YKTG", "YKVN", "PLHN", "SLDO", "0", "0"],
    ["RFGT1000", "Yuki Refrigerator", "YKTG", "YKVN", "PLHN", "SLDO", "0", "0"],
    ["ABCD1000", "Yuki Test Product", "YKTG", "YKVN", "PLSG", "SLDO", "0", "0"],
    ["WSMC1000", "Yuki Washing Machine", "HAWA", "YKVN", "PLHN", "SLDO", "0", "0"],
    ["DRYR1000", "Yuki Dryer 1000", "YKTG", "YKVN", "PLSG", "SLEX", "0", "0"],
    ["FRZR1000", "Yuki Freezer", "YKTG", "YKVN", "PLSG", "SLEX", "0", "0"],
    ["RFGT1000", "Yuki Refrigerator", "YKTG", "YKVN", "PLSG", "SLEX", "0", "0"],
    ["WSMC1000", "Yuki Washing Machine", "HAWA", "YKVN", "PLSG", "SLEX", "0", "0"],
    ["ACDT1000", "Yuki Air Conditioner 1000", "YKTG", "YKVN", "PLHN", "SLHV", "0", "10.00"],
    ["DRYR1000", "Yuki Dryer 1000", "YKTG", "YKVN", "PLHN", "SLHV", "0", "0"],
    ["FRZR1000", "Yuki Freezer", "YKTG", "YKVN", "PLHN", "SLHV", "0", "0"],
    ["RFGT1000", "Yuki Refrigerator", "YKTG", "YKVN", "PLHN", "SLHV", "0", "0"],
    ["WSMC1000", "Yuki Washing Machine", "HAWA", "YKVN", "PLHN", "SLHV", "0", "0"],
  ],
  detailedRows: [
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDGYuki%20Dryer%201000YKTGDRYR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDGYuki%20Dryer%201000YKTGDRYR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDG",
      Maktx: "Yuki Dryer 1000",
      Mtart: "YKTG",
      Matnr: "DRYR1000",
      Matnr_T: "Yuki Dryer 1000",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDGYuki%20Dryer%201000YKTGDRYR1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDGYuki%20Dryer%201000YKTGDRYR1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDG",
      Maktx: "Yuki Dryer 1000",
      Mtart: "YKTG",
      Matnr: "DRYR1000",
      Matnr_T: "Yuki Dryer 1000",
      Bwkey: "PLSG",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.36.32_4.0.0.0.4.12.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDGYuki%20FreezerYKTGFRZR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.36.32_4.0.0.0.4.12.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDGYuki%20FreezerYKTGFRZR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDG",
      Maktx: "Yuki Freezer",
      Mtart: "YKTG",
      Matnr: "FRZR1000",
      Matnr_T: "Yuki Freezer",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.41.32_4.0.0.0.4.17.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDGYuki%20RefrigeratorYKTGRFGT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.41.32_4.0.0.0.4.17.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDGYuki%20RefrigeratorYKTGRFGT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDG",
      Maktx: "Yuki Refrigerator",
      Mtart: "YKTG",
      Matnr: "RFGT1000",
      Matnr_T: "Yuki Refrigerator",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.44.32_4.0.0.0.4.20.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDGYuki%20Washing%20MachineHAWAWSMC1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.44.32_4.0.0.0.4.20.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDGYuki%20Washing%20MachineHAWAWSMC1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDG",
      Maktx: "Yuki Washing Machine",
      Mtart: "HAWA",
      Matnr: "WSMC1000",
      Matnr_T: "Yuki Washing Machine",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Dryer%201000YKTGDRYR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Dryer%201000YKTGDRYR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDO",
      Maktx: "Yuki Dryer 1000",
      Mtart: "YKTG",
      Matnr: "DRYR1000",
      Matnr_T: "Yuki Dryer 1000",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Dryer%201000YKTGDRYR1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Dryer%201000YKTGDRYR1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDO",
      Maktx: "Yuki Dryer 1000",
      Mtart: "YKTG",
      Matnr: "DRYR1000",
      Matnr_T: "Yuki Dryer 1000",
      Bwkey: "PLSG",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.36.32_4.0.0.0.4.12.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20FreezerYKTGFRZR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.36.32_4.0.0.0.4.12.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20FreezerYKTGFRZR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDO",
      Maktx: "Yuki Freezer",
      Mtart: "YKTG",
      Matnr: "FRZR1000",
      Matnr_T: "Yuki Freezer",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.41.32_4.0.0.0.4.17.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20RefrigeratorYKTGRFGT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.41.32_4.0.0.0.4.17.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20RefrigeratorYKTGRFGT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDO",
      Maktx: "Yuki Refrigerator",
      Mtart: "YKTG",
      Matnr: "RFGT1000",
      Matnr_T: "Yuki Refrigerator",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.41.32_4.0.0.0.4.17.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Test%20ProductYKTGABCD1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.41.32_4.0.0.0.4.17.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Test%20ProductYKTGABCD1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDO",
      Maktx: "Yuki Test Product",
      Mtart: "YKTG",
      Matnr: "ABCD1000",
      Matnr_T: "Yuki Test Product",
      Bwkey: "PLSG",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.44.32_4.0.0.0.4.20.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Washing%20MachineHAWAWSMC1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.44.32_4.0.0.0.4.20.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Washing%20MachineHAWAWSMC1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDO",
      Maktx: "Yuki Washing Machine",
      Mtart: "HAWA",
      Matnr: "WSMC1000",
      Matnr_T: "Yuki Washing Machine",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLEXYuki%20Dryer%201000YKTGDRYR1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLEXYuki%20Dryer%201000YKTGDRYR1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLEX",
      Maktx: "Yuki Dryer 1000",
      Mtart: "YKTG",
      Matnr: "DRYR1000",
      Matnr_T: "Yuki Dryer 1000",
      Bwkey: "PLSG",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.36.32_4.0.0.0.4.12.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLEXYuki%20FreezerYKTGFRZR1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.36.32_4.0.0.0.4.12.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLEXYuki%20FreezerYKTGFRZR1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLEX",
      Maktx: "Yuki Freezer",
      Mtart: "YKTG",
      Matnr: "FRZR1000",
      Matnr_T: "Yuki Freezer",
      Bwkey: "PLSG",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.41.32_4.0.0.0.4.17.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLEXYuki%20RefrigeratorYKTGRFGT1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.41.32_4.0.0.0.4.17.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLEXYuki%20RefrigeratorYKTGRFGT1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLEX",
      Maktx: "Yuki Refrigerator",
      Mtart: "YKTG",
      Matnr: "RFGT1000",
      Matnr_T: "Yuki Refrigerator",
      Bwkey: "PLSG",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.44.32_4.0.0.0.4.20.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLEXYuki%20Washing%20MachineHAWAWSMC1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.44.32_4.0.0.0.4.20.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLEXYuki%20Washing%20MachineHAWAWSMC1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLEX",
      Maktx: "Yuki Washing Machine",
      Mtart: "HAWA",
      Matnr: "WSMC1000",
      Matnr_T: "Yuki Washing Machine",
      Bwkey: "PLSG",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.49.32_4.0.0.0.4.25.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLHVYuki%20Air%20Conditioner%201000YKTGACDT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.49.32_4.0.0.0.4.25.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLHVYuki%20Air%20Conditioner%201000YKTGACDT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLHV",
      Maktx: "Yuki Air Conditioner 1000",
      Mtart: "YKTG",
      Matnr: "ACDT1000",
      Matnr_T: "Yuki Air Conditioner 1000",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "10.00",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLHVYuki%20Dryer%201000YKTGDRYR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.39.32_4.0.0.0.4.15.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLHVYuki%20Dryer%201000YKTGDRYR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLHV",
      Maktx: "Yuki Dryer 1000",
      Mtart: "YKTG",
      Matnr: "DRYR1000",
      Matnr_T: "Yuki Dryer 1000",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.36.32_4.0.0.0.4.12.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLHVYuki%20FreezerYKTGFRZR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.36.32_4.0.0.0.4.12.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLHVYuki%20FreezerYKTGFRZR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLHV",
      Maktx: "Yuki Freezer",
      Mtart: "YKTG",
      Matnr: "FRZR1000",
      Matnr_T: "Yuki Freezer",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.41.32_4.0.0.0.4.17.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLHVYuki%20RefrigeratorYKTGRFGT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.41.32_4.0.0.0.4.17.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLHVYuki%20RefrigeratorYKTGRFGT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLHV",
      Maktx: "Yuki Refrigerator",
      Mtart: "YKTG",
      Matnr: "RFGT1000",
      Matnr_T: "Yuki Refrigerator",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.44.32_4.0.0.0.4.20.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLHVYuki%20Washing%20MachineHAWAWSMC1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.44.44.32_4.0.0.0.4.20.0.0.4.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLHVYuki%20Washing%20MachineHAWAWSMC1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLHV",
      Maktx: "Yuki Washing Machine",
      Mtart: "HAWA",
      Matnr: "WSMC1000",
      Matnr_T: "Yuki Washing Machine",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.39.32_4.0.0.0.4.15.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLREYuki%20Dryer%201000YKTGDRYR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.39.32_4.0.0.0.4.15.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLREYuki%20Dryer%201000YKTGDRYR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLRE",
      Maktx: "Yuki Dryer 1000",
      Mtart: "YKTG",
      Matnr: "DRYR1000",
      Matnr_T: "Yuki Dryer 1000",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.36.32_4.0.0.0.4.12.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLREYuki%20FreezerYKTGFRZR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.36.32_4.0.0.0.4.12.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLREYuki%20FreezerYKTGFRZR1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLRE",
      Maktx: "Yuki Freezer",
      Mtart: "YKTG",
      Matnr: "FRZR1000",
      Matnr_T: "Yuki Freezer",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.41.32_4.0.0.0.4.17.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLREYuki%20RefrigeratorYKTGRFGT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.41.32_4.0.0.0.4.17.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLREYuki%20RefrigeratorYKTGRFGT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLRE",
      Maktx: "Yuki Refrigerator",
      Mtart: "YKTG",
      Matnr: "RFGT1000",
      Matnr_T: "Yuki Refrigerator",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.44.32_4.0.0.0.4.20.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLREYuki%20Washing%20MachineHAWAWSMC1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.44.32_4.0.0.0.4.20.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLREYuki%20Washing%20MachineHAWAWSMC1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLRE",
      Maktx: "Yuki Washing Machine",
      Mtart: "HAWA",
      Matnr: "WSMC1000",
      Matnr_T: "Yuki Washing Machine",
      Bwkey: "PLHN",
      Labst: "0",
      Labst_E: "PC",
      Bdmng: "0",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.49.32_4.0.0.0.4.25.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Air%20Conditioner%201000YKTGACDT1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.49.32_4.0.0.0.4.25.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Air%20Conditioner%201000YKTGACDT1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDO",
      Maktx: "Yuki Air Conditioner 1000",
      Mtart: "YKTG",
      Matnr: "ACDT1000",
      Matnr_T: "Yuki Air Conditioner 1000",
      Bwkey: "PLSG",
      Labst: "253.00",
      Labst_E: "PC",
      Bdmng: "10.00",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.49.32_4.0.0.0.4.25.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLEXYuki%20Air%20Conditioner%201000YKTGACDT1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.49.32_4.0.0.0.4.25.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLEXYuki%20Air%20Conditioner%201000YKTGACDT1000PLSG_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLEX",
      Maktx: "Yuki Air Conditioner 1000",
      Mtart: "YKTG",
      Matnr: "ACDT1000",
      Matnr_T: "Yuki Air Conditioner 1000",
      Bwkey: "PLSG",
      Labst: "871.00",
      Labst_E: "PC",
      Bdmng: "10.00",
      Bdmng_E: "PC",
    },
    {
      __metadata: {
        id: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.49.32_4.0.0.0.4.25.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Air%20Conditioner%201000YKTGACDT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        uri: "http://s35.gb.ucc.cit.tum.de/sap/opu/odata/sap/Z_QUERY_MATERIALSTOCK_CDS/Z_QUERY_MATERIALSTOCK('V2.46.49.32_4.0.0.0.4.25.0.0.4.0.8.4.X.1.0.1.0.0.0.0.0.0.0_YKVNSLDOYuki%20Air%20Conditioner%201000YKTGACDT1000PLHN_0GDKTC9EXA4QEYAEAKKJG0F04OQ35CTB')",
        type: "Z_QUERY_MATERIALSTOCK_CDS.Z_QUERY_MATERIALSTOCKType",
      },
      Bukrs: "YKVN",
      Lgort: "SLDO",
      Maktx: "Yuki Air Conditioner 1000",
      Mtart: "YKTG",
      Matnr: "ACDT1000",
      Matnr_T: "Yuki Air Conditioner 1000",
      Bwkey: "PLHN",
      Labst: "1705.00",
      Labst_E: "PC",
      Bdmng: "10.00",
      Bdmng_E: "PC",
    },
  ],
};

onClickPdf();

async function onClickPdf() {
  console.log("onClickPdf");

  const pngUrl = "https://i.ibb.co/L621CHV/2002.png";
  const pngImageBytes = await fetch(pngUrl).then((res) => res.arrayBuffer());

  const pdfDoc = await PDFLib.PDFDocument.create();

  const pngImage = await pdfDoc.embedPng(pngImageBytes);

  await processPdf(data, pdfDoc, pngImage);

  const pdfBytes = await pdfDoc.save();

  // Save the PDF file to the filesystem
  fs.writeFileSync("document2.pdf", pdfBytes);

  // Remove the link from the document
  //   console.log(pdfBytes);
}

async function processPdf(data, pdfDoc, pngImage) {
  const fontBold = await pdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);
  const fontNormal = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);

  const pngDims = pngImage.scale(0.2);

  const year = "2024";
  const companyName = "Yuki Vietnam";
  const user = "LEARN-033";
  const currentDate = new Date().toLocaleDateString();

  let firstPage = true;

  // Group detailedRows by Material and Plant
  const groupedData = data.detailedRows.reduce((acc, row) => {
    const material = row.Matnr;
    const plant = row.Bwkey;
    const key = `${material}_${plant}`;

    if (!acc[key]) {
      acc[key] = [];
    }

    acc[key].push(row);
    return acc;
  }, {});

  // Iterate over each group
  Object.keys(groupedData).forEach((key) => {
    const detailedRows = groupedData[key];
    const firstRow = detailedRows[0];

    const page = pdfDoc.addPage(PDFLib.PageSizes.A4);
    var yPosition = { value: PDFLib.PageSizes.A4[1] - 20 };

    page.drawImage(pngImage, {
      x: 1,
      y: page.getHeight() - pngDims.height + 15,
      width: pngDims.width,
      height: pngDims.height,
    });

    initialFormatting(page, year, companyName, user, currentDate, yPosition, fontBold, fontNormal, firstPage);

    const headerData = requiredHeadersData.reduce((accumulator, headerItem) => {
      const columnKey = headerItem.column;
      const columnName = headerItem.columnName;

      if (firstRow.hasOwnProperty(columnKey)) {
        const dataValue = firstRow[columnKey];
        accumulator[columnName] = dataValue;
      }

      return accumulator;
    }, {});

    drawHeaderData(page, headerData, yPosition, fontBold, fontNormal);

    const measuresData = detailedRows.map((row) => {
      return requiredMeasuresData.reduce((accumulator, measureItem) => {
        const columnKey = measureItem.column;
        const columnName = measureItem.columnName;

        if (row.hasOwnProperty(columnKey)) {
          const dataValue = row[columnKey];
          accumulator[columnName] = dataValue;
        }

        return accumulator;
      }, {});
    });
    console.log(measuresData);

    drawMeasuresTable(page, measuresData, yPosition, fontBold, fontNormal);

    firstPage = false; // Set to false after processing the first page
  });
}

function drawHeaderData(page, headerData, yPosition, fontBold, fontNormal) {
  const xStart = 50;
  const lineHeight = 30;

  for (const [key, value] of Object.entries(headerData)) {
    const keyText = `${key}: `;
    const keyFontSize = 14;
    const valueFontSize = 16;

    // Draw the key text with fontBold
    page.drawText(keyText, {
      x: xStart,
      y: yPosition.value,
      size: keyFontSize,
      font: fontBold,
    });

    // Calculate the width of the key text to position the value
    const keyTextWidth = fontBold.widthOfTextAtSize(keyText, keyFontSize);

    // Draw the value text with fontNormal
    page.drawText(`${value}`, {
      x: xStart + keyTextWidth,
      y: yPosition.value,
      size: valueFontSize,
      font: fontNormal,
    });

    yPosition.value -= lineHeight;
  }
  yPosition.value -= 20;
}

function drawMeasuresTable(page, measuresData, yPosition, fontBold, fontNormal) {
  const xStart = 50;
  const lineHeight = 20;
  const columnWidth = 140;

  page.drawText("Stock Details", {
    x: xStart,
    y: yPosition.value,
    size: 20,
    font: fontBold,
  });

  yPosition.value -= 30;

  // Get the headers from measuresData
  const headers = Object.keys(measuresData[0]);

  // Draw table headers
  headers.forEach((header, index) => {
    page.drawText(header, {
      x: xStart + index * columnWidth,
      y: yPosition.value,
      size: 14,
      font: fontBold,
    });
  });

  yPosition.value -= lineHeight;

  // Draw table rows
  measuresData.forEach((row) => {
    headers.forEach((header, index) => {
      const value = row[header] || "";
      page.drawText(value.toString(), {
        x: xStart + index * columnWidth,
        y: yPosition.value,
        size: 14,
        font: fontNormal,
      });
    });
    yPosition.value -= lineHeight;
  });

  // Calculate totals for measure columns
  const totals = headers.reduce((acc, header) => {
    acc[header] = measuresData.reduce((sum, row) => {
      const value = parseFloat(row[header]);
      return sum + (isNaN(value) ? 0 : value);
    }, 0);
    return acc;
  }, {});

  yPosition.value -= lineHeight;

  // Draw totals row
  headers.forEach((header, index) => {
    const value = header === "Storage Location" ? "Total" : totals[header].toFixed(2);
    page.drawText(value, {
      x: xStart + index * columnWidth,
      y: yPosition.value,
      size: 14,
      font: fontBold,
    });
  });
}

function initialFormatting(page, year, companyName, user, currentDate, yPosition, fontBold, fontNormal, firstPage) {
  //Initial Formatting
  // Embed fonts

  let rightSideAlignX = page.getWidth() - 200;
  let rightSideColumnHeight = 15;

  // Add 'Year: 2024' (left-aligned, normal text)
  page.drawText(`Fiscal Year: ${year}`, {
    x: rightSideAlignX,
    y: yPosition.value,
    size: 12,
    font: fontNormal,
  });
  yPosition.value -= rightSideColumnHeight;

  // Add 'Company Name: Yuki' (left-aligned, normal text)
  page.drawText(`Company Name: ${companyName}`, {
    x: rightSideAlignX,
    y: yPosition.value,
    size: 12,
    font: fontNormal,
  });
  yPosition.value -= rightSideColumnHeight;

  // Add 'User: LEARN-032' (left-aligned, normal text)
  page.drawText(`User: ${user}`, {
    x: rightSideAlignX,
    y: yPosition.value,
    size: 12,
    font: fontNormal,
  });
  yPosition.value -= rightSideColumnHeight;

  // Add 'Date: ' with current date (left-aligned, normal text)
  page.drawText(`Date: ${currentDate}`, {
    x: rightSideAlignX,
    y: yPosition.value,
    size: 12,
    font: fontNormal,
  });
  yPosition.value -= 30;

  if (firstPage) {
    const title = "Available-To-Promise Report";
    const titleFontSize = 24;
    const titleWidth = fontBold.widthOfTextAtSize(title, titleFontSize);

    page.drawText(title, {
      x: (page.getWidth() - titleWidth) / 2,
      y: yPosition.value,
      size: titleFontSize,
      font: fontBold,
    });

    yPosition.value -= 50;
  }

  page.drawText("Material Details", {
    x: 50,
    y: yPosition.value,
    size: 20,
    font: fontBold,
  });

  yPosition.value -= 30;
}
