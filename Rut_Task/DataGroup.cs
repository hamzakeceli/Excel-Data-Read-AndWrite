namespace Rut_Task

{
    public struct DataGroup
    {

        public string T_GK_1 { get; set; }
        public string T_GK_2 { get; set; }
        public string T_GK_3 { get; set; }
        public string T_GK_4 { get; set; }
        public string T_GK_5 { get; set; }

        public DataGroup(string TGK1, string TGK2, string TGK3, string TGK4, string TGK5)
        {

            T_GK_1 = TGK1;
            T_GK_2 = TGK2;
            T_GK_3 = TGK3;
            T_GK_4 = TGK4;
            T_GK_5 = TGK5;

        }
    }
}
