namespace ExileExcel.Common
{

    public enum PrintPageSize
    {
        // data sourse:
        // http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.pagesetup.aspx

        /// <summary>
        /// Letter paper (8.5 in. by 11 in.)
        /// </summary>
        USLetter = 1,

        /// <summary>
        /// Letter small paper (8.5 in. by 11 in.)
        /// </summary>
        USLetterSmall = 2,

        /// <summary>
        /// Tabloid paper (11 in. by 17 in.)
        /// </summary>
        USTabloid = 3,

        /// <summary>
        /// Ledger paper (17 in. by 11 in.)
        /// </summary>
        Ledger = 4,

        /// <summary>
        /// Legal paper (8.5 in. by 14 in.)
        /// </summary>
        Legal = 5,

        /// <summary>
        /// Statement paper (5.5 in. by 8.5 in.)
        /// </summary>
        Statement = 6,

        /// <summary>
        /// Executive paper (7.25 in. by 10.5 in.)
        /// </summary>
        Executive = 7,

        /// <summary>
        ///  A3 paper (297 mm by 420 mm)
        /// </summary>
        A3,

        /// <summary>
        /// A4 paper (210 mm by 297 mm)
        /// </summary>
        A4 = 9,

        /// <summary>
        /// A4 small paper (210 mm by 297 mm)
        /// </summary>
        A4Small = 10,

        /// <summary>
        /// A5 paper (148 mm by 210 mm)
        /// </summary>
        A5 = 11,

        /// <summary>
        /// B4 paper (250 mm by 353 mm)
        /// </summary>
        B4JIS = 12,

        /// <summary>
        /// B5 paper (176 mm by 250 mm)
        /// </summary>
        B5JIS = 13,

        /// <summary>
        ///  Folio paper (8.5 in. by 13 in.)
        /// </summary>
        Folio = 14,

        /// <summary>
        /// Quarto paper (215 mm by 275 mm)
        /// </summary>
        Quarto = 15,

        /// <summary>
        /// Standard paper (10 in. by 14 in.)
        /// </summary>
        Standard10X14 = 16,

        /// <summary>
        /// Standard paper (11 in. by 17 in.)
        /// </summary>
        Standard11X17 = 17,

        /// <summary>
        /// Note paper (8.5 in. by 11 in.)
        /// </summary>
        Note = 18,

        /// <summary>
        /// #9 envelope (3.875 in. by 8.875 in.)
        /// </summary>
        Envelope9 = 19,

        /// <summary>
        /// #10 envelope (4.125 in. by 9.5 in.)
        /// </summary>
        Envelope10 = 20,

        /// <summary>
        /// #11 envelope (4.5 in. by 10.375 in.)
        /// </summary>
        Envelope11 = 21,

        /// <summary>
        /// #12 envelope (4.75 in. by 11 in.)
        /// </summary>
        Envelope12 = 22,

        /// <summary>
        /// #14 envelope (5 in. by 11.5 in.)
        /// </summary>
        Envelope14 = 23,

        /// <summary>
        /// C paper (17 in. by 22 in.)
        /// </summary>
        CPaper = 24,

        /// <summary>
        /// D paper (22 in. by 34 in.)
        /// </summary>
        DPaper = 25,

        /// <summary>
        /// E paper (34 in. by 44 in.)
        /// </summary>
        EPaper = 26,

        /// <summary>
        /// DL envelope (110 mm by 220 mm)
        /// </summary>
        EnvelopeDL = 27,

        /// <summary>
        /// C5 envelope (162 mm by 229 mm)
        /// </summary>
        EnvelopeC5 = 28,

        /// <summary>
        /// C3 envelope (324 mm by 458 mm)
        /// </summary>
        EnvelopeC3 = 29,

        /// <summary>
        /// C4 envelope (229 mm by 324 mm)
        /// </summary>
        EnvelopeC4 = 30,

        /// <summary>
        /// C6 envelope (114 mm by 162 mm)
        /// </summary>
        EnvelopeC6 = 31,

        /// <summary>
        /// C65 envelope (114 mm by 229 mm)
        /// </summary>
        EnvelopeC65 = 32,

        /// <summary>
        /// B4 envelope (250 mm by 353 mm)
        /// </summary>
        EnvelopeB4 = 33,

        /// <summary>
        /// B5 envelope (176 mm by 250 mm)
        /// </summary>
        EnvelopeB5 = 34,

        /// <summary>
        /// B6 envelope (176 mm by 125 mm)
        /// </summary>
        EnvelopeB6 = 35,

        /// <summary>
        /// Italy envelope (110 mm by 230 mm)
        /// </summary>
        EnvelopeItaly = 36,

        /// <summary>
        /// Monarch envelope (3.875 in. by 7.5 in.).
        /// </summary>
        EnvelopeMonarch = 37,

        /// <summary>
        /// 6 3/4 envelope (3.625 in. by 6.5 in.)
        /// </summary>
        Envelope6 = 38,

        /// <summary>
        /// US standard fanfold (14.875 in. by 11 in.)
        /// </summary>
        USFanfold = 39,

        /// <summary>
        /// German standard fanfold (8.5 in. by 12 in.)
        /// </summary>
        GermanStandardFanfold = 40,

        /// <summary>
        /// German legal fanfold (8.5 in. by 13 in.)
        /// </summary>
        GermanLegalFanfold = 41,

        /// <summary>
        /// B4 (250 mm by 353 mm)
        /// </summary>
        B4 = 42,

        /// <summary>
        /// Japanese double postcard (200 mm by 148 mm)
        /// </summary>
        JapaneseDoublePostcard = 43,

        /// <summary>
        /// Standard paper (9 in. by 11 in.)
        /// </summary>
        Paper9 = 44,

        /// <summary>
        /// Standard paper (10 in. by 11 in.)
        /// </summary>
        Paper10 = 45,

        /// <summary>
        /// Standard paper (15 in. by 11 in.)
        /// </summary>
        Paper15 = 46,

        /// <summary>
        /// Invite envelope (220 mm by 220 mm)
        /// </summary>
        EnvelopeInvite = 47,

        /// <summary>
        /// Letter extra paper (9.275 in. by 12 in.)
        /// </summary>
        PaperExtraLetter = 50,

        /// <summary>
        /// Legal extra paper (9.275 in. by 15 in.)
        /// </summary>
        PaperExtraLegal = 51,

        /// <summary>
        /// Tabloid extra paper (11.69 in. by 18 in.)
        /// </summary>
        PaperExtraTabloid = 52,

        /// <summary>
        /// A4 extra paper (236 mm by 322 mm)
        /// </summary>
        PaperA4Extra = 53,

        /// <summary>
        /// Letter transverse paper (8.275 in. by 11 in.)
        /// </summary>
        PaperLetterTransverse = 54,

        /// <summary>
        /// A4 transverse paper (210 mm by 297 mm)
        /// </summary>
        PaperA4Transverse = 55,

        /// <summary>
        /// Letter extra transverse paper (9.275 in. by 12 in.)
        /// </summary>
        PaperLetterExtraTransverse = 56,

        /// <summary>
        /// SuperA/SuperA/A4 paper (227 mm by 356 mm)
        /// </summary>
        SuperA = 57,

        /// <summary>
        /// SuperB/SuperB/A3 paper (305 mm by 487 mm)
        /// </summary>
        SuperB = 58,

        /// <summary>
        /// Letter plus paper (8.5 in. by 12.69 in.)
        /// </summary>
        PaperLetterPlus = 59,

        /// <summary>
        /// A4 plus paper (210 mm by 330 mm)
        /// </summary>
        PaperA4Plus = 60,

        /// <summary>
        /// A5 transverse paper (148 mm by 210 mm)
        /// </summary>
        PaperA5Tranves = 61,

        /// <summary>
        /// JIS B5 transverse paper (182 mm by 257 mm)
        /// </summary>
        PaperJISB5Transverse = 62,

        /// <summary>
        /// A3 extra paper (322 mm by 445 mm)
        /// </summary>
        PaperA3Extra = 63,

        /// <summary>
        /// A5 extra paper (174 mm by 235 mm)
        /// </summary>
        PaperA5Extra = 64,

        /// <summary>
        /// B5 extra paper (201 mm by 276 mm)
        /// </summary>
        PaperB5Extra = 65,
        /// <summary>
        /// A2 paper (420 mm by 594 mm)
        /// </summary>
        PaperA2 = 66,

        /// <summary>
        /// A3 transverse paper (297 mm by 420 mm)
        /// </summary>
        PaperA3Transverse = 67,

        /// <summary>
        /// A3 extra transverse paper (322 mm by 445 mm)
        /// </summary>
        PaperA3ExtraTransverse = 68
    }
}
