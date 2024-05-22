using ExcelDna.Integration;

public static class Timefunctions
{
    public static object GetCurrentTime()
    {
        return XlCall.RTD("RtdFunctions.TimeServer", null, "NOW");
    }
}
