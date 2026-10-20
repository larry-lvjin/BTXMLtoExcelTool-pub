using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace BtFiller2
{
    public static class ExcelFiller
    {
        private static readonly string TemplateB64 =
        "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" +
        "AAAMAAwAJgMAAPgyAAAAAA==";

        public static void Fill(string xmlPath, string outputPath)
        {
            // 实现细节未公开（算法核心）
            throw new NotImplementedException("该方法的实现不在开源范围内");
        }

        // ── 行号写入表 ────────────────────────────────────────────────
        private static List<(int row, bool isPct, double val)> BuildWrites(TestData d)
        {
            // 实现细节未公开（算法核心）
            throw new NotImplementedException("该方法的实现不在开源范围内");
        }

        // ── 写入单元格（保留 s 属性/样式）────────────────────────────
        private static void WriteCell(
            XDocument doc, XNamespace ns, Dictionary<int, XElement> rowCache,
            int rowIdx, string col, double value)
        {
            // 实现细节未公开（算法核心）
            throw new NotImplementedException("该方法的实现不在开源范围内");
        }

        // ── 修复 0% → 0.00%（numFmtId 9 → 10）───────────────────────
        private static void FixPctStyle(
            XDocument sheetDoc, XNamespace ns, XDocument stylesDoc, int rowIdx, string col)
        {
            // 实现细节未公开（算法核心）
            throw new NotImplementedException("该方法的实现不在开源范围内");
        }

        private static int ColNum(string cref)
        {
            // 实现细节未公开（算法核心）
            throw new NotImplementedException("该方法的实现不在开源范围内");
        }

        private static byte[] ToBytes(XDocument doc) =>
            Encoding.UTF8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                doc.Root!.ToString());

        private static string FindSheetKey(Dictionary<string, byte[]> entries)
        {
            // 实现细节未公开（算法核心）
            throw new NotImplementedException("该方法的实现不在开源范围内");
        }

        private static TestData ExtractData(string xmlPath)
        {
            // 实现细节未公开（算法核心）
            throw new NotImplementedException("该方法的实现不在开源范围内");
        }

        // BV-10-C "Measurement at Max Power" — 18 values:
        // per channel (Ch0, Ch39, Ch78) collect 6 readings in document order:
        //   GFSK Power 2-DH5, DPSK Power 2-DH5, DPSK-GFSK 2-DH5,
        //   GFSK Power 3-DH5, DPSK Power 3-DH5, DPSK-GFSK 3-DH5
        private static List<double> ExtractMaxPower10(XElement root)
        {
            // 实现细节未公开（算法核心）
            throw new NotImplementedException("该方法的实现不在开源范围内");
        }

        private static List<double> WorstPerGroup(XElement root, string tpId)
        {
            // 实现细节未公开（算法核心）
            throw new NotImplementedException("该方法的实现不在开源范围内");
        }

        private class TestData
        {
            public List<double> Avg1DH5=new(), MaxPower10=new(), Mod=new(),
                                FreqAccu=new(), Drift=new(), DriftRate=new(),
                                Edr=new(), SpurBR=new(), SpurEdr=new();
        }
    }
}
