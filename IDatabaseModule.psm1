# IDatabaseModule.psm1

# Define the IDatabase interface using Add-Type
Add-Type -TypeDefinition @"
namespace DatabaseNamespace
{
    public interface IDatabase
    {
        void OpenConnection();
        void CloseConnection();
        void ExecuteDml();
    }

    public interface ITransactionalDatabase
    {
        void CommitTransaction();
        void RollbackTransaction();
    }
}
"@
