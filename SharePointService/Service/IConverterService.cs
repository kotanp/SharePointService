namespace SharePointService.Service
{
    public interface IConverterService
    {
        /*
         * Converts the document byte array to pdf then returns it as an byte array
         */
        public string ConvertToPdf(byte[] docItem, string fileExtension);

        /*
         * Returns the doc's revision counter 
         */
        public string GetDocRevision(byte[] docItem, string fileExtension);
    }
}
