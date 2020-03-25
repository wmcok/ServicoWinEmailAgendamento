Imports System.Data.Entity
Imports System.Data.Entity.ModelConfiguration.Conventions

Public Class DbRelatorios
    Inherits DbContext

    Public Property Emails As DbSet(Of Email)

    Public Shared Function Create() As DbRelatorios
        Return New DbRelatorios()
    End Function
    Protected Overrides Sub OnModelCreating(ByVal modelBuilder As DbModelBuilder)
        modelBuilder.Conventions.Remove(Of PluralizingTableNameConvention)()
        modelBuilder.Conventions.Remove(Of OneToManyCascadeDeleteConvention)()
        modelBuilder.Conventions.Remove(Of ManyToManyCascadeDeleteConvention)()
        MyBase.OnModelCreating(modelBuilder)
    End Sub

End Class