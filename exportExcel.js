import ExcelJS from "exceljs";
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

onClick();

async function onClick() {
  //initialize worksheet
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("My Sheet", {
    pageSetup: { paperSize: 9, orientation: "landscape" },
  });
  const pngUrl = "https://i.ibb.co/L621CHV/2002.png";
  const pngImageBytes = await fetch(pngUrl).then((res) => res.arrayBuffer());
  const imageId = workbook.addImage({
    buffer: pngImageBytes,
    extension: "png",
  });

  processExcel(worksheet, data, imageId);

  //download the file
  await workbook.xlsx.writeFile("Available-to-Promise_Report.xlsx");
}

async function processExcel(worksheet, data, imageId) {
  worksheet.addImage(imageId, {
    tl: { col: 0, row: 0 },
    br: { col: 2, row: 7 },
    // ext: { width: 120, height: 120 },
  });
  const year = "2024";
  const companyName = "Yuki Vietnam";
  const user = "LEARN-033";
  const currentDate = new Date().toLocaleDateString();

  // Determine the rightmost column letter based on the number of columns
  const totalColumns = data.columns.length;
  const startColumnIndex = 3; // Column 'C' is index 3
  const rightColumnIndex = startColumnIndex + totalColumns - 1; // Adjusted calculation
  const rightColumnLetter = String.fromCharCode(64 + rightColumnIndex);

  // Merge cells for the title dynamically
  worksheet.mergeCells(`C1:${rightColumnLetter}7`);
  worksheet.getCell("C1").value = "Available-to-Promise Report";
  worksheet.getCell("C1").alignment = { vertical: "middle", horizontal: "center" };
  worksheet.getCell("C1").font = { size: 24, bold: true };

  // Set B9 "Company Name"
  worksheet.getCell("B9").value = "Company Name";

  // Merge C9 to middle column and set companyName
  const middleColumnIndex = Math.floor((startColumnIndex + rightColumnIndex) / 2);
  const middleColumnLetter = String.fromCharCode(64 + middleColumnIndex);
  worksheet.mergeCells(`C9:${middleColumnLetter}9`);
  worksheet.getCell("C9").value = companyName;

  // Set B10 "Fiscal Year"
  worksheet.getCell("B10").value = "Fiscal Year";
  // Merge C10 to middle column and set year
  worksheet.mergeCells(`C10:${middleColumnLetter}10`);
  worksheet.getCell("C10").value = year;

  // Set the positions for "User" and "Date"
  const userStartColumnIndex = middleColumnIndex + 1;
  const userStartColumnLetter = String.fromCharCode(64 + userStartColumnIndex);

  worksheet.getCell(`${userStartColumnLetter}9`).value = "User";
  // Merge next columns for 'User' value
  const userValueStartIndex = userStartColumnIndex + 1;
  const userValueStartLetter = String.fromCharCode(64 + userValueStartIndex);
  worksheet.mergeCells(`${userValueStartLetter}9:${rightColumnLetter}9`);
  worksheet.getCell(`${userValueStartLetter}9`).value = user;

  // Set 'Date'
  worksheet.getCell(`${userStartColumnLetter}10`).value = "Date";
  // Merge next columns for 'Date' value
  worksheet.mergeCells(`${userValueStartLetter}10:${rightColumnLetter}10`);
  worksheet.getCell(`${userValueStartLetter}10`).value = currentDate;

  // Move to B12 for data columns and rows
  const startRow = 12;

  // Set column headers starting from column B
  data.columns.forEach((col, index) => {
    const columnLetter = String.fromCharCode(66 + index); // 66 is 'B'
    const cellAddress = `${columnLetter}${startRow}`;
    worksheet.getCell(cellAddress).value = col.header;
    worksheet.getColumn(columnLetter).width = col.width;
  });

  // Add data rows starting from startRow + 1
  data.rows.forEach((rowData, rowIndex) => {
    rowData.forEach((value, colIndex) => {
      const columnLetter = String.fromCharCode(66 + colIndex); // Start from 'B'
      const cellAddress = `${columnLetter}${startRow + 1 + rowIndex}`;
      worksheet.getCell(cellAddress).value = value;
    });
  });

  // //add all those data to the actual worksheet
  // worksheet.columns = data.columns;
  // worksheet.addRows(data.rows);
  const row = worksheet.lastrow;
}
