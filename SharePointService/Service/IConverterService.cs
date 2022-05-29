namespace SharePointService.Service
{
    public interface IConverterService
    {
        /*
         * Converts the document byte array to pdf then returns it as an byte array
         */
        public string ConvertToPdf(byte[] docItem, string fileExtension);
    }
}
