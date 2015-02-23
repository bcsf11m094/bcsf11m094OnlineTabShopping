Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Reflection
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Diagnostics.CodeAnalysis
Imports System.Security
Imports System.Security.Permissions

''' <summary>Base class of all custom enumeration.</summary>
<SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix", Justification:="Because it is a base class for all enums, the suffix ""Enum"" is just fine.")> _
Public MustInherit Class CustomEnum

    Public MustOverride ReadOnly Property Name As String
    Public MustOverride ReadOnly Property Index As Int32

End Class



''' <summary>
''' Base-class for valueless generic custom enums that have to overwrite the combination class and base class for
''' all valued generic custom enums (ValueEnum).
''' </summary>
''' <typeparam name="TEnum">The type of your subclass</typeparam><remarks>
''' Hints for implementors: You must ensure that only one instance of each enum-value exists. This is easily reached by
''' declaring the constructor(s) private, sealing the class and exposing the enum-values as static fields. If you are
''' implementing them through static getter properties, make sure lazy initialization is used and that always the same
''' instance is returned with every call.
''' </remarks>
<SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
<SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix", Justification:="Because it is a base class for all enums, the suffix ""Enum"" is just fine.")> _
<SuppressMessage("Microsoft.Design", "CA1046:DoNotOverloadOperatorEqualsOnReferenceTypes", Justification:="CustomEnum behaves like a value type.")> _
<DebuggerDisplay("{DebuggerDisplayValue}")> _
Public MustInherit Class CustomEnum(Of TEnum As {CustomEnum(Of TEnum, TCombi)}, TCombi As {CustomEnum(Of TEnum, TCombi).Combi, New})
    Inherits CustomEnum

    'Private Fields

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared _Members As TEnum()
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared _CaseSensitive As Boolean?
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared _NameComparer As IEqualityComparer(Of String)
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared _IsFirstInstance As Boolean = True
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared _Lock As New Object()
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private _Name As String
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private _Index As Int32 = -1
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private _Combi As TCombi = Nothing

    'Constructors

    ''' <summary>Called by implementors to create a new instance of TEnum (when assigning the instance to a static field). 
    ''' Important: Make your constructors private to ensure there are no instances except the ones initialized 
    ''' by your subclass! Null values are not supported and throw an ArgumentNullException.</summary>
    ''' <exception cref="ArgumentNullException">An ArgumentNullException is thrown if aValue is null.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this instance's type is not of type <typeparam name="TEnum" />.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException may be thrown (check only in full-trust possible) if this instance has non-private constructors.</exception>
    Protected Sub New()
        Me.New(Nothing)
    End Sub

    ''' <summary>Called by implementors to create a new instance of TEnum (when assigning the instance to a static field). 
    ''' Important: Make your constructors private to ensure there are no instances except the ones initialized 
    ''' by your subclass! Null values are not supported and throw an ArgumentNullException.</summary>
    ''' <param name="caseSensitive">Leave null for automatic determination (recommended), or set explicitely.</param>
    ''' <exception cref="ArgumentNullException">An ArgumentNullException is thrown if aValue is null.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this instance's type is not of type <typeparam name="TEnum" />.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException may be thrown (check only in full-trust possible) if this instance has non-private constructors.</exception>
    Protected Sub New(ByVal caseSensitive As Boolean?)
        'Make sure no evil cross subclass is changing our static variable
        If (Not GetType(TEnum).IsAssignableFrom(Me.GetType())) Then
            Throw New InvalidOperationException("Internal error in " & Me.GetType().Name & "! Change the first type parameter from """ & GetType(TEnum).Name & """ to """ & Me.GetType().Name & """.")
        End If
        'Make sure only the first subclass is affecting our static variables
        If (_IsFirstInstance) Then
            _IsFirstInstance = False
            'Check constructors
            CheckConstructors()
            'Assign static variables
            _CaseSensitive = caseSensitive
        End If
    End Sub

    'Public Properties

    ''' <summary>Returns the name of this member (the name of the static field or static getter property this member 
    ''' was assigned to in the subclass). Watch out: Do not access from within the subclasses constructor!</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this property is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is also thrown if this instance is not 
    ''' assigned to a public static readonly field or property of TEnum.</exception>
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="readonly")> _
    Public Overrides ReadOnly Property Name As String
        Get
            'Check whether the name is already assigned
            Dim myResult As String = _Name
            If (myResult Is Nothing) Then
                InitAndReturnMembers()
                myResult = _Name
                If (myResult Is Nothing) Then
                    Throw New InvalidOperationException("Detached instance error! Ensure that all members are assigned to a public static readonly field or property.")
                End If
            End If
            'Return the name
            Return myResult
        End Get
    End Property

    ''' <summary>Returns the zero-based index position of this member. Watch out: Do not access from within the subclasses constructor!
    ''' If there are static fields as well as static getter properties, the fields have the lower index. The order is the same as it is 
    ''' returnd from Type.GetFields() and Type.GetProperties() and should correspond to the order the fields/properties have been declared. 
    ''' The index may be used by to compare members.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this property is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is also thrown if this instance is not 
    ''' assigned to a public static readonly field or property of TEnum.</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="readonly")> _
    Public Overrides ReadOnly Property Index As Int32
        Get
            'Check whether the index is already assigned
            Dim myResult As Int32 = _Index
            If (myResult < 0) Then
                InitAndReturnMembers()
                myResult = _Index
                If (myResult < 0) Then
                    Throw New InvalidOperationException("Detached instance error! Ensure that all members are assigned to a public static readonly field or property.")
                End If
            End If
            'Return the index
            Return myResult
        End Get
    End Property

    'Public Methods

    ''' <summary>Returns true if one of the given members equals this member, false otherwise. If the given paramarray 
    ''' is null, false is returned. If the array contains null values they are ignored.</summary>
    Public Function [In](ByVal ParamArray members As TEnum()) As Boolean
        'Check the args
        If (members Is Nothing) OrElse (members.Length = 0) Then Return False
        'Loop through given members
        Dim myInstance As TEnum = DirectCast(Me, TEnum)
        For Each myMember As TEnum In members
            If (myMember Is Nothing) Then Continue For
            If (myMember.Equals(myInstance)) Then Return True
        Next
        'Otherwise return false
        Return False
    End Function

    ''' <summary>Returns true if the given combination contains this member, false otherwise. False is also returned if
    ''' the combination is null.</summary>
    ''' <param name="combination">The combination to check (null is allowed).</param>
    Public Function [In](ByVal combination As TCombi) As Boolean
        'Check args
        If (combination Is Nothing) Then Return False
        'Determine whether it contains us
        Dim myInstance As TEnum = DirectCast(Me, TEnum)
        Return combination.Contains(myInstance)
    End Function

    ''' <summary>Returns the member with the next higher index. Watch out: Do not access from within the subclasses constructor!
    ''' If this member is the last one, then it depends on parameter <paramref name="loop" /> whether the first member is returned 
    ''' (<paramref name="loop" /> is set to <c>true</c>) or null (<paramref name="loop" /> is set to <c>false</c>). If the enum does 
    ''' not contain any members, null is returned.</summary>
    ''' <param name="loop">Whether the first member should be returned if the end is reached (true) or null (false).</param>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this method is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this method is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Function GetNext(ByVal [loop] As Boolean) As TEnum
        Dim myIndex As Int32 = Index + 1
        Dim myMembers As TEnum() = Members
        If (myIndex >= myMembers.Length) Then
            If ([loop]) Then Return myMembers(0)
            Return Nothing
        End If
        Return myMembers(myIndex)
    End Function

    ''' <summary>Returns the member with the next lower index. If this member is the first one, then it depends on parameter 
    ''' <paramref name="loop" /> whether the last member is returned (<paramref name="loop" /> is set to <c>true</c>) or null 
    ''' (<paramref name="loop" /> is set to <c>false</c>). If the enum does not contain any members, null is returned.</summary>
    ''' <param name="loop">Whether the last member should be returned if the start is reached (true) or null (false).</param>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this method is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this method is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Function GetPrevious(ByVal [loop] As Boolean) As TEnum
        Dim myIndex As Int32 = Index - 1
        Dim myMembers As TEnum() = Members
        If (myIndex < 0) Then
            If ([loop]) Then Return myMembers(myMembers.Length - 1)
            Return Nothing
        End If
        Return myMembers(myIndex)
    End Function

    'Public Class Properties

    ''' <summary>Returns the type of the combination class.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Public Shared ReadOnly Property CombiType As Type
        Get
            Return GetType(TCombi)
        End Get
    End Property

    ''' <summary>Returns whether the names passed to function <see cref="GetMemberByName" /> are treated case sensitive
    ''' or not (using <see cref="StringComparer.Ordinal" /> resp. <see cref="StringComparer.OrdinalIgnoreCase" />).
    ''' The default behavior is that they are case-insensitive except there would be two or more entries that would
    ''' cause an ambiguity.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException may be thrown if this property is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this property is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared ReadOnly Property CaseSensitive As Boolean
        Get
            Dim myResult As Boolean? = _CaseSensitive
            If (myResult Is Nothing) Then
                Return InitAndReturnCaseSensitive(Members)
            End If
            Return myResult.Value
        End Get
    End Property

    ''' <summary>Gets the first defined member of this enum (the one with index 0). If the enum does not contain any members,
    ''' null is returned.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this property is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this property is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Public Shared ReadOnly Property First As TEnum
        Get
            Dim myMembers As TEnum() = Members
            If (myMembers.Length = 0) Then Return Nothing
            Return myMembers(0)
        End Get
    End Property

    ''' <summary>Gets the last defined member of this enum (the one with the highest index). If the enum does not contain any members,
    ''' null is returned.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this property is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this property is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Public Shared ReadOnly Property Last As TEnum
        Get
            Dim myMembers As TEnum() = Members
            If (myMembers.Length = 0) Then Return Nothing
            Return myMembers(myMembers.Length - 1)
        End Get
    End Property

    'Public Class Functions

    ''' <summary>Returns an empty combination (a reference to <see cref="Combi.Empty">Combi.Empty</see>).-</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi() As TCombi
        Return Combi.Empty
    End Function

    ''' <summary>Returns a combination containing the member (if the member is null, a reference to 
    ''' <see cref="Combi.Empty">Combi.Empty</see> is returned, otherwise <see cref="ToCombi">member.ToCombi()</see>).</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi(ByVal member As TEnum) As TCombi
        Return Combi.GetInstanceOptimized(member)
    End Function

    ''' <summary>Returns a combination containing the given members.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi(ByVal member1 As TEnum, ByVal ParamArray member2 As TEnum()) As TCombi
        Return Combi.GetInstanceOptimized(member1, member2)
    End Function

    ''' <summary>Returns and instance that contains the given members. If members is null or empty, 
    ''' <see cref="Combi.Empty">Combi.Empty</see> is returned, if the member contains exactly one non-null-member,
    ''' <see cref="ToCombi">members[0].ToCombi()</see> is returned, otherwise a new instance containing 
    ''' the given members.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi(ByVal members As TEnum()) As TCombi
        Return Combi.GetInstanceOptimized(members)
    End Function

    ''' <summary>Returns and instance that contains the given members. If members is null or empty, 
    ''' <see cref="Combi.Empty">Combi.Empty</see> is returned, otherwise a reference to the given member.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi(ByVal members As Combi) As TCombi
        Return Combi.GetInstanceOptimized(members)
    End Function

    ''' <summary>Constructs a new combination and initializes the provided members. Watch out: This function is
    ''' not thread-safe, use synchronization means to prevent other threads from manipulating the collection
    ''' until the combination is returned.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi(ByVal memberCollection As IEnumerable(Of TEnum)) As TCombi
        If (memberCollection Is Nothing) Then Return Combi.Empty
        Dim myArray As TEnum() = ToArray(Of TEnum)(memberCollection)
        Return Combi.GetInstanceOptimized(myArray)
    End Function

    ''' <summary>Returns a new array containing all members defined by this enum (in index order).</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMembers() As TEnum()
        Dim myMembers As TEnum() = Members
        Dim myResult(myMembers.Length - 1) As TEnum
        Array.Copy(myMembers, myResult, myMembers.Length)
        Return myResult
    End Function

    ''' <summary>Returns the names of all members defined by this enum (in index order). The names are the ones of
    ''' the "public static fields" and "public static getter properties without indexer" of this CustomEnum's type.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetNames() As String()
        Dim myMembers As TEnum() = Members
        Dim myResult(myMembers.Length - 1) As String
        For i As Int32 = 0 To myMembers.Length - 1
            myResult(i) = myMembers(i)._Name 'speed optimized (property is always initialized)
        Next
        Return myResult
    End Function

    ''' <summary>
    ''' Returns the member of the given name or null if not found. Property <see cref="CaseSensitive" /> tells whether 
    ''' <paramref name="name" /> is treated case-sensitive or not. If name is null, an ArgumentNullException is thrown. 
    ''' If the subclass is incorrectly implemented and has duplicate names defined, an InvalidOperationException is thrown. 
    ''' This function is thread-safe.
    ''' </summary>
    ''' <param name="name">The name to look up.</param>
    ''' <returns>The enum entry or null if not found.</returns>
    ''' <remarks>A full duplicate check is performed the first time this method (or GetMembersByName) is called.</remarks>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMemberByName(ByVal name As String) As TEnum
        'Check the argument
        If (name Is Nothing) Then Throw New ArgumentNullException("name")
        'Get/initialize members and comparer
        Dim myMembers As TEnum() = Members
        Dim myComparer As IEqualityComparer(Of String) = NameComparer
        'Return the first name found (it is always unique, ensured when the NameComparer is initialized)
        For Each myMember As TEnum In myMembers
            If (myComparer.Equals(myMember._Name, name)) Then Return myMember 'speed optimized, _Name is always initialized
        Next
        'Otherwise return null
        Return Nothing
    End Function

    ''' <summary>
    ''' Returns the member of the given name or null if not found using the given comparer
    ''' to perform the comparison. An ArgumentException is thrown if the result would be ambiguous
    ''' according to the given comparer. If there are no special reasons don't use this method but the
    ''' one without the comparer overload as it is optimized to perform the duplicate check only
    ''' once and not every time the method is used. This method is thread-safe if the nameComparer
    ''' is thread-safe (or not being manipulated during the method call).</summary>
    ''' <param name="name">The name to look up.</param>
    ''' <param name="nameComparer">The comparer to use for the equality comparison of the strings (null defaults to StringComparer.Ordinal).</param>
    ''' <returns>The member or null if not found (or throws an exception if more than one is found).</returns>
    ''' <exception cref="ArgumentNullException">An ArgumentNullException is thrown if name is null.</exception>
    ''' <exception cref="ArgumentException">An ArgumentException is thrown if the result would be ambiguous according to the given comparer.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMemberByName(ByVal name As String, ByVal nameComparer As IEqualityComparer(Of String)) As TEnum
        'Check the argument
        If (name Is Nothing) Then Throw New ArgumentNullException("name")
        'Use optimized method if possible
        If (nameComparer Is Nothing) Then nameComparer = StringComparer.Ordinal
        If (Object.Equals(nameComparer, _NameComparer)) Then Return GetMemberByName(name)
        'Get the members
        Dim myMembers As TEnum() = Members
        'Get the first found member but continue looping
        Dim myResult As TEnum = Nothing
        For Each myMember As TEnum In myMembers
            If (nameComparer.Equals(myMember._Name, name)) Then
                If (myResult Is Nothing) Then
                    myResult = myMember
                Else
                    Throw New ArgumentException("According to the given comparer at least two ambiguous matches were found!")
                End If
            End If
        Next
        'Return the result 
        Return myResult
    End Function

    ''' <summary>
    ''' Returns the found members with the given names or Combi.Empty if not found. Property <see cref="CaseSensitive" /> 
    ''' tells whether the parameters are treated case-sensitive or not. For easier consumation by other languages two parameters
    ''' are used instead of one. If both arguments are null, empty or contain only unknown members, Combi.Empty is returned. 
    ''' Duplicates and null values within the array are ignored. If the subclass is incorrectly implemented and has duplicate names 
    ''' defined, an InvalidOperationException is thrown. This function is thread-safe.
    ''' </summary>
    ''' <param name="name">The name to look up.</param>
    ''' <param name="additionalNames">The name to look up.</param>
    ''' <returns>The found members as a combination.</returns>
    ''' <remarks>A full duplicate check is performed the first time this method (or GetMemberByName) is called.</remarks>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId:="ByNames", Justification:="The spelling is okay like this.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMembersByNames(ByVal name As String, ByVal ParamArray additionalNames As String()) As TCombi
        'Handle single value
        If (additionalNames Is Nothing) OrElse (additionalNames.Length = 0) Then
            'Handle null/empty
            If (name Is Nothing) Then Return Combi.Empty
            'Get member
            Dim myMember As TEnum = GetMemberByName(name)
            If (myMember Is Nothing) Then Return Combi.Empty
            Return myMember.ToCombi()
        End If
        'Get/initialize comparer and dictionary
        Dim myComparer As IEqualityComparer(Of String) = NameComparer
        Dim myDict As New Dictionary(Of String, String)(additionalNames.Length + 1, myComparer)
        'Fill in names
        If (name IsNot Nothing) Then myDict.Add(name, name)
        For Each myName As String In additionalNames
            If (myName Is Nothing) Then Continue For
            myDict.Item(myName) = myName
        Next
        'Return the result
        Return GetMembersByNames(myDict)
    End Function

    ''' <summary>
    ''' Returns the found members with the given names or Combi.Empty if not found. Property <see cref="CaseSensitive" /> 
    ''' tells whether the names are treated case-sensitive or not. If the enumeration is null, empty or contains only unknown 
    ''' members, Combi.Empty is returned. Duplicates and null values within the collection are ignored. If the subclass is 
    ''' incorrectly implemented and has duplicate names defined, an InvalidOperationException is thrown. This function is 
    ''' thread-safe only if the given nameCollection is thread-safe (or not manipulated during the time this function executes).
    ''' </summary>
    ''' <param name="nameCollection">The names to look up.</param>
    ''' <returns>The found members as a combination.</returns>
    ''' <remarks>A full duplicate check is performed the first time this method (or GetMemberByName) is called.</remarks>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId:="ByNames", Justification:="The spelling is okay like this.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMembersByNames(ByVal nameCollection As IEnumerable(Of String)) As TCombi
        'Handle null
        If (nameCollection Is Nothing) Then Return Combi.Empty
        'Get/initialize comparer and dictionary
        Dim myComparer As IEqualityComparer(Of String) = NameComparer
        Dim myDict As Dictionary(Of String, String) = Nothing
        Dim myCollection As ICollection(Of String) = TryCast(nameCollection, ICollection(Of String))
        If (myCollection Is Nothing) Then
            myDict = New Dictionary(Of String, String)(myComparer)
        Else
            myDict = New Dictionary(Of String, String)(myCollection.Count, myComparer)
        End If
        'Fill in names
        For Each myName As String In nameCollection
            If (myName Is Nothing) Then Continue For
            myDict.Item(myName) = myName
        Next
        'Return the result
        Return GetMembersByNames(myDict)
    End Function

    ''' <summary>
    ''' Returns the members of the given names. If the a name is not found or null, it is ignored. If there are ambiguous matches according to the
    ''' given comparer, an ArgumentException is thrown. If nameComparer is null, the comparison is performed using an ordinal comparer. If there are 
    ''' no special reasons don't use this method but the one without the comparer overload as it is optimized to perform the duplicate check only
    ''' once and not every time the method is called. This method is thread-safe if the nameComparer is thread-safe (or not being manipulated during 
    ''' the method call).</summary>
    ''' <param name="nameCollection">The names to look up.</param>
    ''' <param name="nameComparer">The comparer to use for the equality comparison of the strings (null defaults to StringComparer.Ordinal).</param>
    ''' <returns>The members found as a combination (an ArgumentException is thrown if there are duplicates).</returns>
    ''' <exception cref="ArgumentException">An ArgumentException is thrown if the result would be ambiguous according to the given comparer.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId:="ByNames", Justification:="The spelling is okay like this.")> _
    <SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId:="0", Justification:="Null values are allowed.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMembersByNames(ByVal nameCollection As IEnumerable(Of String), ByVal nameComparer As IEqualityComparer(Of String)) As TCombi
        'Handle null
        If (nameCollection Is Nothing) Then Return Combi.Empty
        'Get/initialize comparer and dictionary
        If (nameComparer Is Nothing) Then nameComparer = StringComparer.Ordinal
        Dim myDict As Dictionary(Of String, String) = Nothing
        Dim myCollection As ICollection(Of String) = TryCast(nameCollection, ICollection(Of String))
        If (myCollection Is Nothing) Then
            myDict = New Dictionary(Of String, String)(nameComparer)
        Else
            myDict = New Dictionary(Of String, String)(myCollection.Count, nameComparer)
        End If
        'Fill in names
        For Each myName As String In nameCollection
            If (myName Is Nothing) Then Continue For
            myDict.Item(myName) = myName
        Next
        'Return the result
        Return GetMembersByNames(myDict)
    End Function

    ''' <summary>
    ''' Gets the member by index. If the index is out-of-bounds null is returned. This function is faster
    ''' that calling GetMembers() with the index because function GetMembers has to copy the array and this
    ''' function does not need to.
    ''' </summary>
    ''' <param name="index">The zero-based index of the member.</param>
    ''' <returns>The member at the given zero-based index (null if out-of-range)</returns>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMemberByIndex(ByVal index As Int32) As TEnum
        If (index < 0) Then Return Nothing
        Dim myMembers As TEnum() = Members
        If (index >= myMembers.Length) Then Return Nothing
        Return myMembers(index)
    End Function

    ''' <summary>
    ''' Gets the members by the given indices. If some of the indices are out-of-bounds, they are ignored. If the array is null or
    ''' empty, it is ignored. This function is faster that calling GetMembers() with the index because function GetMembers has to 
    ''' copy the array and this function does not need to.
    ''' </summary>
    ''' <param name="index">The zero-based index of the first member to look up.</param>
    ''' <param name="additionalIndexes">Additional indices to look up (split into two parameters for better consumation by C#).</param>
    ''' <returns>A combination with the members of the given indices set.</returns>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMembersByIndexes(ByVal index As Int32, ByVal ParamArray additionalIndexes As Int32()) As TCombi
        Return Combi.GetInstanceByIndexesOptimized(index, additionalIndexes)
    End Function

    ''' <summary>
    ''' Returns the found members with the given indices. If the enumerable is null, empty or contains only indices that are
    ''' out-of-range, Combi.Empty is returned. Duplicate values within the collection are ignored. If the subclass is 
    ''' incorrectly implemented and has duplicate names defined, an InvalidOperationException is thrown. This function is 
    ''' thread-safe only if the given indexCollection is thread-safe (or not manipulated during the time this function executes).
    ''' </summary>
    ''' <param name="indexCollection">The indices to look up.</param>
    ''' <returns>The found members as a combination.</returns>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMembersByIndexes(ByVal indexCollection As IEnumerable(Of Int32)) As TCombi
        Dim myArray As Int32() = ToArray(Of Int32)(indexCollection)
        Return Combi.GetInstanceByIndexesOptimized(myArray)
    End Function

    'Public Operators

    ''' <summary>Combines two enum values to a Combination using a binary OR operation. It is more efficient to use 
    ''' <see cref="GetCombi">GetCombi(..)</see> to combine more than two members. There is no 
    ''' difference between the "OR" and the "+" operators. This operation is thread-safe.</summary>
    ''' <param name="arg1">The first member to combine (null is valid).</param>
    ''' <param name="arg2">The second member to combine (null is valid).</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator Or(ByVal arg1 As CustomEnum(Of TEnum, TCombi), ByVal arg2 As TEnum) As TCombi
        'Handle null
        If (arg1 Is Nothing) Then
            If (arg2 Is Nothing) Then Return Combi.Empty
            Return arg2.ToCombi()
        End If
        If (arg2 Is Nothing) Then Return arg1.ToCombi()
        'Return new combination
        Dim myArg1 As TEnum = CType(arg1, TEnum)
        Return Combi.GetInstanceOptimized(New TEnum() {myArg1, arg2})
    End Operator

    ''' <summary>Combines two enum values through XOR and returns a new combination instance of the two. If you are combining
    ''' more than two values it is more efficient to initialize an empty combination and the call the Toggle method
    ''' subsequently (or ToggleRange) because each XOR operation allocates a new Combination instance.</summary>
    ''' <param name="arg1">The first member to combine (null is valid).</param>
    ''' <param name="arg2">The second member to combine (null is valid).</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator Xor(ByVal arg1 As CustomEnum(Of TEnum, TCombi), ByVal arg2 As TEnum) As TCombi
        'Handle null
        If (arg1 Is Nothing) Then
            If (arg2 Is Nothing) Then Return Combi.Empty
            Return arg2.ToCombi()
        End If
        If (arg2 Is Nothing) Then Return arg1.ToCombi()
        'If the members are equal, return an empty combination
        If (arg1 = arg2) Then Return Combi.Empty
        'Otherwise return normal or
        Return (arg1 Or arg2)
    End Operator

    ''' <summary>Binare AND operation of two single members. It returns arg1 if it equals arg2, null otherwise. This operation is thread-safe.</summary>
    ''' <param name="arg1">The first argument, a member (null is valid).</param>
    ''' <param name="arg2">The second argument, a member (null is valid).</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator And(ByVal arg1 As CustomEnum(Of TEnum, TCombi), ByVal arg2 As TEnum) As TEnum
        'If the members are equal, return a member
        If (arg1 = arg2) Then Return CType(arg1, TEnum)
        'Otherwise return null
        Return Nothing
    End Operator

    ''' <summary>Returns a new instance of Combination that contains all members that are not in arg. This operation is thread-safe.</summary>
    ''' <param name="arg">The argument, a combination of members.</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator Not(ByVal arg As CustomEnum(Of TEnum, TCombi)) As TCombi
        'Handle null
        If (arg Is Nothing) Then Return Combi.All
        'Return inverted result
        Dim myResult As TCombi = Combi.All
        Return (myResult - CType(arg, TEnum))
    End Operator

    ''' <summary>Binare AND operation of a single member and a combination of members. It returns arg1 if it is contained in arg2, null otherwise. 
    ''' This operation is thread-safe.</summary>
    ''' <param name="arg1">The first argument, a member.</param>
    ''' <param name="arg2">The second argument, a combination of members.</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator And(ByVal arg1 As CustomEnum(Of TEnum, TCombi), ByVal arg2 As TCombi) As TEnum
        'Handle null
        If (arg1 Is Nothing) OrElse (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return Nothing
        'Merge both flags
        Dim myArg1 As TEnum = DirectCast(arg1, TEnum)
        If (arg2.Contains(myArg1)) Then Return myArg1
        Return Nothing
    End Operator

    ''' <summary>Binare AND operation of a single member and a combination of members. It returns arg2 if it is contained in arg1, null otherwise. 
    ''' This operation is thread-safe.</summary>
    ''' <param name="arg1">The first argument, a combination of members.</param>
    ''' <param name="arg2">The second argument, a member.</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator And(ByVal arg1 As TCombi, ByVal arg2 As CustomEnum(Of TEnum, TCombi)) As TEnum
        Return (arg2 And arg1)
    End Operator

    ''' <summary>Subtracts a member from a member (null is returned if arg2 equlas arg1, otherwise arg1 is returned).
    ''' This operation is thread-safe.</summary>
    ''' <param name="arg1">The first argument, a member.</param>
    ''' <param name="arg2">The argument to subtract, a member.</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator -(ByVal arg1 As CustomEnum(Of TEnum, TCombi), ByVal arg2 As TEnum) As TEnum
        If (arg1 Is Nothing) Then Return Nothing
        If (arg2 Is Nothing) Then Return CType(arg1, TEnum)
        If (arg1.Equals(arg2)) Then Return Nothing
        Return CType(arg1, TEnum)
    End Operator

    ''' <summary>Subtracts multiple members from a member (null is returned if arg2 contains arg1, otherwise arg1 is returned).
    ''' This operation is thread-safe.</summary>
    ''' <param name="arg1">The first argument, a member.</param>
    ''' <param name="arg2">The argument to subtract, a combination of members.</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator -(ByVal arg1 As CustomEnum(Of TEnum, TCombi), ByVal arg2 As TCombi) As TEnum
        'Handle null
        If (arg1 Is Nothing) Then Return Nothing
        Dim myArg1 As TEnum = CType(arg1, TEnum)
        If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return myArg1
        'Handle combination
        If (arg2.Contains(myArg1)) Then
            Return Nothing
        End If
        Return myArg1
    End Operator

    ''' <summary>Combines two enum values to a Combination using a binary OR operation. It is more efficient to use 
    ''' <see cref="GetCombi">GetCombi(..)</see> to combine more than two members. There is no 
    ''' difference between the "+" and the "OR" operators. This operation is thread-safe.</summary>
    ''' <param name="arg1">The first member to combine (null is valid).</param>
    ''' <param name="arg2">The second member to combine (null is valid).</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator +(ByVal arg1 As CustomEnum(Of TEnum, TCombi), ByVal arg2 As TEnum) As TCombi
        Return (arg1 Or arg2)
    End Operator

    ''' <summary>Compares two member and returns true if they are equal. This operation is thead-safe.</summary>
    ''' <param name="arg1">The first argument, a member (null is valid).</param>
    ''' <param name="arg2">The second argument, a member (null is valid).</param>
    Public Shared Operator =(ByVal arg1 As CustomEnum(Of TEnum, TCombi), ByVal arg2 As TEnum) As Boolean
        If (Object.ReferenceEquals(arg1, arg2)) Then Return True 'if both are null, true is returned
        If (arg1 Is Nothing) OrElse (arg2 Is Nothing) Then Return False
        Return arg1.Equals(arg2)
    End Operator

    ''' <summary>Compares two member and returns true if they are inequal.</summary>
    ''' <param name="arg1">The first argument, a member (null is valid).</param>
    ''' <param name="arg2">The second argument, a member (null is valid).</param>
    Public Shared Operator <>(ByVal arg1 As CustomEnum(Of TEnum, TCombi), ByVal arg2 As TEnum) As Boolean
        Return (Not (arg1 = arg2))
    End Operator

    'Framework Properties

    ''' <summary>Returns true if the implementation of this CustomEnum looks correct.</summary>
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="The exception is not needed.")> _
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared ReadOnly Property IsValid() As Boolean
        Get
            'If the initialization does not throw an exception, return true
            Try
                InitAndReturnMembers()
                Return True
            Catch
            End Try
            'Otherwise return false
            Return False
        End Get
    End Property

    'Framework Methods

    ''' <summary>Returns a combination with this member set. This method is optimized and returns always the same instance.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    Public Function ToCombi() As TCombi
        Dim myResult As TCombi = _Combi
        If (myResult Is Nothing) Then
            myResult = Combi.GetInstanceRaw(CType(Me, TEnum))
            _Combi = myResult
        End If
        Return myResult
    End Function

    ''' <summary>Returns the name of the enum.</summary>
    Public Overrides Function ToString() As String
        Dim myResult As String = _Name
        If (myResult Is Nothing) Then
            Try
                Return Name 'initializes the name if possible
            Catch ex As InvalidOperationException
                Return "[not initialized]"
            End Try
        End If
        Return myResult
    End Function

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private ReadOnly Property DebuggerDisplayValue As String
        Get
            Dim myResult As String = _Name
            If (myResult Is Nothing) Then
                Try
                    Return Name 'initializes the name if possible
                Catch ex As InvalidOperationException
                    Return "[unknown]"
                End Try
            End If
            Return myResult
        End Get
    End Property

    ''' <summary>Returns true if the given object is a member and equals this member (see according Equals overload), 
    ''' or if it is of type Combination and contains only this instance as member (see according Equals overload).</summary>
    ''' <param name="obj">The other instance.</param>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        'Handle null
        If (obj Is Nothing) Then Return False
        'Handle Member
        Dim myMember As TEnum = TryCast(obj, TEnum)
        If (myMember IsNot Nothing) Then
            Return Equals(myMember)
        End If
        'Handle Combination
        Dim myCombination As Combi = TryCast(obj, Combi)
        If (myCombination IsNot Nothing) Then
            Return Equals(myCombination)
        End If
        'Otherwise return false
        Return False
    End Function

    ''' <summary>By default, returns true if the other instance is the same reference as this one, false otherwise. 
    ''' This behavior may be overwritten in the subclass, eg. if there are two defined values that are equal.</summary>
    ''' <param name="other">The other instance.</param>
    Public Overridable Overloads Function Equals(ByVal other As TEnum) As Boolean
        If (Object.ReferenceEquals(Me, other)) Then Return True
        Return False
    End Function

    ''' <summary>Returns true if the combination contains this instance as the only member.</summary>
    ''' <param name="other">The other instance.</param>
    Public Overloads Function Equals(ByVal other As Combi) As Boolean
        If (other Is Nothing) Then Return False
        Return other.Equals(Me)
    End Function

    ''' <summary>Returns the hashcode of the value to ensure member and value have the same hashcodes.</summary>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Overrides Function GetHashCode() As Int32
        Return Index
    End Function

    'Private Properties

    ''' <summary>Returns all members of this enum. The first time this property is called they are evaluated through reflection 
    ''' and then cached in a static variable. Watch out: Do not call this function from within the subclasses constructor!</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if a property or method is accessed from 
    ''' within the subclasses constructor that calls this property und would cause the member array to be initialized (which is not 
    ''' possible to that time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function 
    ''' is called later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared ReadOnly Property Members As TEnum()
        Get
            Dim myResult As TEnum() = _Members
            If (myResult Is Nothing) Then
                myResult = InitAndReturnMembers()
            End If
            Return myResult
        End Get
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared ReadOnly Property NameComparer As IEqualityComparer(Of String)
        Get
            Dim myResult As IEqualityComparer(Of String) = _NameComparer
            If (myResult Is Nothing) Then
                myResult = InitAndReturnNameComparer(Members)
            End If
            Return myResult
        End Get
    End Property

    ''' <summary>For debugging only. Hides the static infos from the instance but still allows to browse it if needed.</summary>
    Private Shared ReadOnly Property Zzzzz As StaticFields
        Get
            Return StaticFields.Singleton
        End Get
    End Property

    'Private Methods

    ''' <summary>Called after the members of this CustomEnum have been initialized. This method does nothing by
    ''' default but may be overwritten in the subclass.</summary>
    Protected Overridable Sub OnMemberInitialized()
    End Sub

    'Private Functions

    Private Shared Function GetMembersByNames(ByVal dict As Dictionary(Of String, String)) As TCombi
        Dim myMembers As TEnum() = Members
        Dim myResult As New List(Of TEnum)(dict.Count)
        'Fill in the members
        For Each myMember As TEnum In myMembers
            If (dict.Remove(myMember._Name)) Then
                myResult.Add(myMember)
                If (dict.Count = 0) Then Exit For
            End If
        Next
        'Return the members as combination
        Select Case myResult.Count
            Case 0
                Return Combi.Empty
            Case 1
                Return myResult(0).ToCombi()
            Case Else
                Return Combi.GetInstanceOptimized(myResult.ToArray())
        End Select
    End Function

    ''' <summary>.NET 2.0 support.</summary>
    Private Shared Function ToArray(Of T)(ByVal collection As IEnumerable(Of T)) As T()
        'Handle null
        If (collection Is Nothing) Then Return New T() {}
        'Handle array
        Dim myArray As T() = TryCast(collection, T())
        If (myArray IsNot Nothing) Then Return myArray
        'Initialize result list
        Dim myResult As List(Of T)
        Dim myCollection As ICollection(Of T) = TryCast(collection, ICollection(Of T))
        If (myCollection Is Nothing) Then
            myResult = New List(Of T)()
        Else
            myResult = New List(Of T)(myCollection.Count)
        End If
        'Fill in elements
        For Each myElement As T In collection
            myResult.Add(myElement)
        Next
        'Return result
        Return myResult.ToArray()
    End Function

    ''' <summary>Initializes the members if they are not yet initialized.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if any property or method
    ''' is accessed from within the subclasses constructor that would cause the member array to be initialized (that
    ''' is not possible to that time because the fields are not yet assigned). An InvalidOperationException is also
    ''' thrown if the same instance is assigned to multiple static fields (that is not possible because Name and Index
    ''' are set to the instance)</exception>
    ''' <returns>A reference to the initialized array.</returns>
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="CustomEnum")> _
    Private Shared Function InitAndReturnMembers() As TEnum()
        Dim myResult As TEnum() = _Members
        If (myResult IsNot Nothing) Then Return myResult
        'Initialize the members
        SyncLock (_Lock)
            myResult = _Members
            If (myResult IsNot Nothing) Then Return myResult
            myResult = PrivateGetMembers()
            If (myResult Is Nothing) Then Throw New InvalidOperationException("Internal error in CustomEnum.")
            'Init comparer and perfom a duplicate check
            If (_NameComparer Is Nothing) Then
                InitAndReturnNameComparer(myResult)
            End If
            'Tell the instance it is initialized
            For Each myMember As TEnum In myResult
                myMember.OnMemberInitialized()
            Next
            'Assign the members (do not assign before the initialization has completed, let other theads wait)
            _Members = myResult
        End SyncLock
        'Return the result
        Return myResult
    End Function

    ''' <summary>Initializes and returns the string comparer used to compare the names.</summary>
    ''' <param name="memberArray">A reference to the member array. Because during initialization it is not yet assigned to _Members it is passed along.</param>
    ''' <exception cref="InvalidOperationException">An invalid operation exception is thrown if there are ambiguous names (not easy to produce in VB).</exception>
    Private Shared Function InitAndReturnNameComparer(ByVal memberArray As TEnum()) As IEqualityComparer(Of String)
        Dim myComparer As IEqualityComparer(Of String) = _NameComparer
        If (myComparer IsNot Nothing) Then Return myComparer
        SyncLock (_Lock)
            myComparer = _NameComparer
            If (myComparer IsNot Nothing) Then Return myComparer
            'Determine the comparer
            myComparer = If(InitAndReturnCaseSensitive(memberArray), StringComparer.Ordinal, StringComparer.OrdinalIgnoreCase)
            'Check for duplicates (happens if the constructor explicitely sets the case-insensitive flag but has two fields that differ only by case,
            'or if the enum has multiple hierarchical subclasses and the overwritten properties do not hide the parent ones (like this is the case 
            'in JScript.NET).
            If (HasDuplicateNames(memberArray, myComparer)) Then
                Throw New InvalidOperationException("Internal error in " & GetType(TEnum).Name & ", the member names are ambiguous.")
            End If
            'If everything is okay, assign the comparer
            _NameComparer = myComparer
            'And return it
            Return myComparer
        End SyncLock
    End Function

    ''' <summary>Gets the members (during initialization).</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if any property or method
    ''' is accessed from within the subclasses constructor that would cause the member array to be initialized (that
    ''' is not possible to that time because the fields are not yet assigned). An InvalidOperationException is also
    ''' thrown if the same instance is assigned to multiple static fields (that is not possible because Name and Index
    ''' are set to the instance)</exception>
    Private Shared Function PrivateGetMembers() As TEnum()
        Dim myList As New List(Of TEnum)
        AddFields(myList)
        AddGetters(myList)
        Return myList.ToArray()
    End Function

    ''' <summary>Adds all public static readonly fields that are of type TEnum (flat, only of class TEnum).</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if any property or method
    ''' is accessed from within the subclasses constructor that would cause the member array to be initialized (that
    ''' is not possible to that time because the fields are not yet assigned). An InvalidOperationException is also
    ''' thrown if the same instance is assigned to multiple static fields (that is not possible because Name and Index
    ''' are set to the instance)</exception>
    Private Shared Sub AddFields(ByVal aList As List(Of TEnum))
        Dim myFlags As BindingFlags = (BindingFlags.Static Or BindingFlags.Public)
        Dim myFields As FieldInfo() = GetType(TEnum).GetFields(myFlags)
        For Each myField As FieldInfo In myFields
            'Ignore read/write fields
            If (Not myField.IsInitOnly) Then Continue For
            'Ignore fields of other types
            If (Not (GetType(TEnum).IsAssignableFrom(myField.FieldType))) Then Continue For
            'Ignore flagged fields
            If (IsFlaggedToIgnore(myField)) Then Continue For
            'Add field
            Dim myEntry As TEnum = CType(myField.GetValue(Nothing), TEnum)
            AddEntry(myEntry, myField.Name, aList)
        Next
    End Sub

    ''' <summary>Adds all public static getter properties without indexer that are of type TEnum (flat, only of class TEnum).</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if any property or method
    ''' is accessed from within the subclasses constructor that would cause the member array to be initialized (that
    ''' is not possible to that time because the fields are not yet assigned). An InvalidOperationException is also
    ''' thrown if the same instance is assigned to multiple static fields (that is not possible because Name and Index
    ''' are set to the instance)</exception>
    Private Shared Sub AddGetters(ByVal aList As List(Of TEnum))
        Dim myFlags As BindingFlags = (BindingFlags.Static Or BindingFlags.Public Or BindingFlags.GetProperty)
        Dim myProperties As PropertyInfo() = GetType(TEnum).GetProperties(myFlags)
        For Each myProperty As PropertyInfo In myProperties
            'Ignore properties of other types
            If (Not (GetType(TEnum).IsAssignableFrom(myProperty.PropertyType))) Then Continue For
            'Look only at read-only properties
            If (myProperty.CanWrite) Then Continue For
            If (Not myProperty.CanRead) Then Continue For
            'Ignore indexed properties
            If (myProperty.GetIndexParameters().Length > 0) Then Continue For
            'Ignore flagged properties
            If (IsFlaggedToIgnore(myProperty)) Then Continue For
            'Invoke the property twice and check whether the same instance is returned (it is a requirement)
            Dim myEntry As TEnum = CType(myProperty.GetValue(Nothing, Nothing), TEnum)
            Dim myEntry2 As TEnum = CType(myProperty.GetValue(Nothing, Nothing), TEnum)
            If (Not Object.ReferenceEquals(myEntry, myEntry2)) Then
                Throw New InvalidOperationException("Internal error in " & GetType(TEnum).Name & "! Property " & myProperty.Name & " returned different instances when invoked multiple times. Ensure always the same instance is returned.")
            End If
            'Add the entry
            AddEntry(myEntry, myProperty.Name, aList)
        Next
    End Sub

    Private Shared Function IsFlaggedToIgnore(ByVal field As FieldInfo) As Boolean
        Return IsFlaggedToIgnore(field.GetCustomAttributes(False))
    End Function

    Private Shared Function IsFlaggedToIgnore(ByVal [property] As PropertyInfo) As Boolean
        Return IsFlaggedToIgnore([property].GetCustomAttributes(False))
    End Function

    Private Shared Function IsFlaggedToIgnore(ByVal attributes As Object()) As Boolean
        If (attributes Is Nothing) OrElse (attributes.Length = 0) Then Return False
        For Each myInstance As Object In attributes
            If (myInstance Is Nothing) Then Continue For
            If (TypeOf myInstance Is CustomEnumIgnoreAttribute) Then Return True
        Next
        Return False
    End Function


    ''' <summary>Adds an entry to the list. During this process it is checked whether the same instance was already added (throws
    ''' an InvalidOperationException) and also assigns the member's name and index.</summary>
    ''' <param name="member">The member to add.</param>
    ''' <param name="name">The field or property name where the member was assigned to.</param>
    ''' <param name="result">A list of members to which the values are added.</param>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if any property or method
    ''' is accessed from within the subclasses constructor that would cause the member array to be initialized (that
    ''' is not possible to that time because the fields are not yet assigned). An InvalidOperationException is also
    ''' thrown if the same instance is assigned to multiple static fields (that is not possible because Name and Index
    ''' are set to the instance)</exception>
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="CustomEnumIgnoreAttribute", Justification:="That's the name of the attribute!")> _
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="readonly")> _
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="OnMemberInitialized")> _
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="GetMemberBy")> _
    Private Shared Sub AddEntry(ByVal member As TEnum, ByVal name As String, ByVal result As List(Of TEnum))
        'Check for instance conflicts
        If (member Is Nothing) Then
            Throw New InvalidOperationException("Internal error in " & GetType(TEnum).Name & "! Do not access any properties (like property Name and Index) nor methods (e.g. all ""GetMemberBy..."") that trigger the initialization of the member array from within the subclasses constructor as the fields are not yet initialized at that time! Also avoid to provide any public static readonly fields of type " & GetType(TEnum).Name & " that have a null-value. You can override method OnMemberInitialized if you need to do some post-initialization, or change the static fields into static getter properties with lazy initialization.")
        End If
        If (member._Name IsNot Nothing) Then
            Throw New InvalidOperationException("Internal error in " & GetType(TEnum).Name & "! It's invalid to assign the same instance to multiple fields/properties that are treated as members (a conflict arises when assigning name and index to the instance). You can use the CustomEnumIgnoreAttribute to flag a field/property to not be treated as a member declaration but as an additional custom field/property of the class.")
        End If
        'Set the name and index
        member._Name = name
        member._Index = result.Count
        'Add to the list
        result.Add(member)
    End Sub

    ''' <summary>Determines whether case-sensitive name comparison is needed (two or more members differ only by name, 
    ''' e.g. when the subclass was written in C#) or not.</summary>
    ''' <param name="memberArray">A reference to the member array. Because during initialization it is not yet assigned to _Members it is passed along.</param>
    ''' <returns>True if case sensitive comparison is needed.</returns>
    Private Shared Function InitAndReturnCaseSensitive(ByVal memberArray As TEnum()) As Boolean
        Dim myResult As Boolean? = _CaseSensitive
        If (myResult Is Nothing) Then
            myResult = HasDuplicateNames(memberArray, StringComparer.OrdinalIgnoreCase)
            _CaseSensitive = myResult
        End If
        Return myResult.Value
    End Function

    ''' <summary>Determines whether there are duplicate names when compared with the given comparer.</summary>
    ''' <param name="memberArray">A reference to the member array. Because during initialization it is not yet assigned to _Members it is passed along.</param>
    ''' <param name="comparer">An equality comparer used to comparer the names.</param>
    ''' <returns>True if there were duplicate names.</returns>
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="The exception is not needed.")> _
    Private Shared Function HasDuplicateNames(ByVal memberArray As TEnum(), ByVal comparer As IEqualityComparer(Of String)) As Boolean
        Dim myDict As New Dictionary(Of String, TEnum)(comparer)
        Try
            For Each myEntry As TEnum In memberArray
                myDict.Add(myEntry._Name, myEntry)
            Next
        Catch
            Return True
        End Try
        Return False
    End Function

    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="MustInherit")> _
    Private Shared Sub CheckConstructors()
        'Check whether we are granted permission to access private members through reflection
        Dim permission As New ReflectionPermission(PermissionState.Unrestricted)
        Try
            permission.Demand()
        Catch ex As SecurityException
            Return
        End Try
        'Get all constructors
        Dim myConstructors As ConstructorInfo() = GetType(TEnum).GetConstructors(BindingFlags.CreateInstance Or BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
        For Each myConstructor As ConstructorInfo In myConstructors
            'Should be private
            If (myConstructor.IsPrivate) Then Continue For
            If (myConstructor.DeclaringType.IsAbstract) Then
                If (myConstructor.IsFamily) Then Continue For
                If (myConstructor.IsFamilyAndAssembly) Then Continue For
            End If
            'Notify implementation error
            Throw New InvalidOperationException("All constructors of " & GetType(TEnum).Name & " must be declared private. If the class is defined as ""MustInherit"", protected constructors are tolerated!")
        Next
    End Sub


    '**********************************************************************
    ' INNER CLASS: StaticInfo
    '**********************************************************************

    ''' <summary>For debugger only. Allowes to hide static information from normal instances but the information is still 
    ''' browsable if needed.</summary>
    <DebuggerDisplay("Expand for static infos...", Name:="(static)")> _
    Private Class StaticFields

        'Private Fields
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private Shared _Singleton As StaticFields

        'Constructors

        Private Sub New()
        End Sub

        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Public Shared ReadOnly Property Singleton As StaticFields
            Get
                Dim myResult As StaticFields = _Singleton
                If (myResult Is Nothing) Then
                    myResult = New StaticFields()
                    _Singleton = myResult
                End If
                Return myResult
            End Get
        End Property

        'Public Properties

        Public Shared ReadOnly Property Members As TEnum()
            Get
                Return CustomEnum(Of TEnum, TCombi).Members
            End Get
        End Property

        Public Shared ReadOnly Property First As TEnum
            Get
                Return CustomEnum(Of TEnum, TCombi).First
            End Get
        End Property

        Public Shared ReadOnly Property Last As TEnum
            Get
                Return CustomEnum(Of TEnum, TCombi).Last
            End Get
        End Property

        Public Shared ReadOnly Property CaseSensitive As Boolean
            Get
                Return CustomEnum(Of TEnum, TCombi).CaseSensitive
            End Get
        End Property

        Public Shared ReadOnly Property IsValid As Boolean
            Get
                Return CustomEnum(Of TEnum, TCombi).IsValid
            End Get
        End Property

        Public Shared ReadOnly Property NameComparer As IEqualityComparer(Of String)
            Get
                Return CustomEnum(Of TEnum, TCombi).NameComparer
            End Get
        End Property

        Public Shared ReadOnly Property CombiType As Type
            Get
                Return CustomEnum(Of TEnum, TCombi).CombiType
            End Get
        End Property

    End Class


    '**************************************************************************
    ' INNER CLASS: Combi
    '**************************************************************************

    ''' <summary>Class that provides support for combining multiple enumeration values, similair to standard enums in combination
    ''' with the <see cref="FlagsAttribute">Flags</see> attribute. A combination is mutually readonly but there are many 
    ''' operators and method that allow to create new combinations. This class aims to be thead-safe. Because of that, it does not 
    ''' support IEnumerable(Of TEnum) directly as arguments for methods and operators. But there is a constructor that takes an 
    ''' IEnumerable(Of TEnum) and like this you can convert your collection into a Combi and then do whatever you like, 
    ''' thead-safe.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix", Justification:="It is no collection and I like it like that.")> _
    <SuppressMessage("Microsoft.Design", "CA1034:NestedTypesShouldNotBeVisible", Justification:="We need very much, that is a feature, not a bug!")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <DebuggerDisplay("{DebuggerDisplayValue}")> _
    Public Class Combi
        Implements IEnumerable(Of TEnum)
        Implements IEquatable(Of Combi)
        Implements IEquatable(Of TEnum)

        'Private Fields
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private Shared _Empty As TCombi = Nothing
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private Shared _All As TCombi = Nothing
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private _Flags As Boolean()
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private _Count As Int32 = -1
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private _HashCode As Int32? = Nothing

        'Constructors

        ''' <summary>Creates a new empty Combi instance. Do not call this constructor, use Combi.Empty if you need 
        ''' a reference to an empty combination, or use an operator overload or a GetCombi-method to get a combination containing
        ''' members.</summary>
        Public Sub New()
        End Sub

        ''' <summary>Returns an instance of Combi containing the given member. This function is thread-safe.</summary>
        <EditorBrowsable(EditorBrowsableState.Never)> _
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Public Shared Function GetInstanceOptimized(ByVal member As TEnum) As TCombi
            If (member Is Nothing) Then Return Empty
            Return member.ToCombi()
        End Function

        ''' <summary>Returns an instance of Combi containing the given members. This function is thread-safe.</summary>
        <EditorBrowsable(EditorBrowsableState.Never)> _
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Public Shared Function GetInstanceOptimized(ByVal member1 As TEnum, ByVal member2 As TEnum()) As TCombi
            'Handle empty
            If ((member2 Is Nothing) OrElse (member2.Length = 0)) Then
                If (member1 Is Nothing) Then Return Empty
                Return member1.ToCombi()
            End If
            'Determine result
            Dim myResult As New TCombi()
            myResult.Set(member1)
            myResult.Set(member2)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Function

        ''' <summary>Returns <see cref="Combi.Empty" /> if null or empty, otherwise an instance containing the given members is returned. 
        ''' This function is thread-safe.</summary>
        ''' <param name="members">The members to get.</param>
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <EditorBrowsable(EditorBrowsableState.Never)> _
        Public Shared Function GetInstanceOptimized(ByVal members As TEnum()) As TCombi
            'Determine result
            Dim myResult As New TCombi()
            myResult.Set(members)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Function

        ''' <summary>Returns <see cref="Combi.Empty" /> if null or empty, otherwise the same instance as provided as argument is returned. This 
        ''' function is thread-safe.</summary>
        ''' <param name="members">The members to get.</param>
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <EditorBrowsable(EditorBrowsableState.Never)> _
        Public Shared Function GetInstanceOptimized(ByVal members As Combi) As TCombi
            If (members Is Nothing) OrElse (members.IsEmpty) Then Return Empty
            Return CType(members, TCombi)
        End Function

        'Watch out: Always return a new instance from GetInstanceRaw-functions, never return the one that was given as a parameter,
        '           don't return Combi.Empty nor Combi.All as the instances may still be manipulated after the call.
        '           This is the only place where this rule applies, in all other places it's save and recommended to recycle 
        '           instances as they are immutable as soon as they leave this class.

        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Private Shared Function GetInstanceRaw() As TCombi
            Return New TCombi()
        End Function

        ''' <summary>Constructs a new combination and initializes the given member.</summary>
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <EditorBrowsable(EditorBrowsableState.Never)> _
        Public Shared Function GetInstanceRaw(ByVal member As TEnum) As TCombi
            Dim myResult As New TCombi()
            myResult.Set(member)
            Return myResult
        End Function

        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Private Shared Function GetInstanceRaw(ByVal members As Combi) As TCombi
            Dim myResult As New TCombi()
            If (members Is Nothing) OrElse (members.Count = 0) Then Return myResult
            Array.Copy(members.Flags, myResult.Flags, members.Flags.Length)
            Return myResult
        End Function


        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <EditorBrowsable(EditorBrowsableState.Never)> _
        Public Shared Function GetInstanceByIndexesOptimized(ByVal indexArray As Int32()) As TCombi
            'Handle empty
            If (indexArray Is Nothing) OrElse (indexArray.Length = 0) Then Return Empty
            'Determine result
            Dim myResult As New TCombi()
            Dim myFlags As Boolean() = myResult.Flags
            Dim myLength As Int32 = myFlags.Length
            For Each myIndex As Int32 In indexArray
                If (myIndex >= myLength) OrElse (myIndex < 0) Then Continue For
                myFlags(myIndex) = True
            Next
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return the result
            Return myResult
        End Function

        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <EditorBrowsable(EditorBrowsableState.Never)> _
        Public Shared Function GetInstanceByIndexesOptimized(ByVal index As Int32, ByVal indexArray As Int32()) As TCombi
            'Determine result
            Dim myResult As New TCombi()
            Dim myFlags As Boolean() = myResult.Flags
            Dim myLength As Int32 = myFlags.Length
            'Add single index
            If (index < myLength) AndAlso (index > -1) Then
                myFlags(index) = True
            End If
            'Add array
            If (indexArray IsNot Nothing) AndAlso (indexArray.Length > 0) Then
                For Each myIndex As Int32 In indexArray
                    If (myIndex >= myLength) OrElse (myIndex < 0) Then Continue For
                    myFlags(myIndex) = True
                Next
            End If
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return the result
            Return myResult
        End Function

        ''' <summary>Returns a combination where all possible members are set.</summary>
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Public Shared ReadOnly Property All() As TCombi
            Get
                Dim myResult As TCombi = _All
                If (myResult Is Nothing) Then
                    'Define the All combination
                    Dim myMembers As TEnum() = Members
                    Select Case myMembers.Length
                        Case 0
                            myResult = Empty
                        Case 1
                            myResult = myMembers(0).ToCombi()
                        Case Else
                            myResult = New TCombi()
                            myResult.SetAll()
                    End Select
                    'Assign the result
                    _All = myResult
                End If
                'Return the result
                Return myResult
            End Get
        End Property

        ''' <summary>Returns an empty combination.</summary>
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Public Shared ReadOnly Property Empty() As TCombi
            Get
                Dim myResult As TCombi = _Empty
                If (myResult Is Nothing) Then
                    myResult = New TCombi()
                    _Empty = myResult
                End If
                Return myResult
            End Get
        End Property

        'Public Properties

        ''' <summary>Gets the member of this combination with the given index. If the index is out-of-range, null is returned.</summary>
        ''' <param name="index">The zero-based index of the member to get.</param>
        Default Public ReadOnly Property Item(ByVal index As Int32) As TEnum
            Get
                'Handle out-of-range
                If (index < 0) Then Return Nothing
                If (index >= Count) Then Return Nothing
                'Find according flag
                Dim myFlags As Boolean() = Flags
                For i As Int32 = 0 To myFlags.Length - 1
                    If (myFlags(i)) Then
                        If (index = 0) Then Return Members(i)
                        index -= 1
                    End If
                Next
                'Should never reach here
                Throw New InvalidOperationException("Internal error in " & Me.GetType().FullName & "!")
            End Get
        End Property

        ''' <summary>Returns the number of members this combination has set. This property is thread-safe.</summary>
        Public ReadOnly Property Count As Int32
            Get
                'Return result from cache
                Dim myResult As Int32 = _Count
                If (myResult > -1) Then Return myResult
                'Initialize result
                Dim myFlags As Boolean() = Flags
                myResult = 0
                For Each myFlag As Boolean In myFlags
                    If (myFlag) Then myResult += 1
                Next
                'Assign result
                _Count = myResult
                'Return result
                Return myResult
            End Get
        End Property

        ''' <summary>Returns true if none of the flags is set, false otherwise. This property is thread-safe.</summary>
        Public ReadOnly Property IsEmpty() As Boolean
            Get
                Dim myCount As Int32 = Count
                Return (myCount = 0)
            End Get
        End Property

        ''' <summary>Returns true if all of the flags are set, false otherwise. If the enum does not define
        ''' any members, true is returned. This property is thread-safe.</summary>
        Public ReadOnly Property IsAllSet() As Boolean
            Get
                'Return true is Count is equal to the number of flags
                Dim myFlags As Boolean() = Flags
                Dim myCount As Int32 = Count
                Return (myFlags.Length = myCount)
            End Get
        End Property

        'Public Methods

        ''' <summary>Determines whether this instance equals another instance (it may be equal to a member or to a combination).</summary>
        ''' <param name="obj">The other instance that is supposed to be equal.</param>
        <EditorBrowsable(EditorBrowsableState.Advanced)> _
        Public Overrides Function Equals(ByVal obj As Object) As Boolean
            'Handle null
            If (obj Is Nothing) Then Return False
            'Handle same reference
            If (Object.ReferenceEquals(Me, obj)) Then Return True
            'Handle Combination
            Dim myCombination As Combi = TryCast(obj, Combi)
            If (myCombination IsNot Nothing) Then
                Return Equals(myCombination)
            End If
            'Handle Member
            Dim myMember As TEnum = TryCast(obj, TEnum)
            If (myMember IsNot Nothing) Then
                Return Equals(myMember)
            End If
            'Otherwise return false
            Return False
        End Function

        ''' <summary>Compares this combination with another one and returns true if they are equal (the same flags are set),
        ''' false otherwise.</summary>
        ''' <param name="other">The members that are supposed to be equal.</param>
        ''' <returns>True if exactly the same members are set, false otherwise.</returns>
        Public Overloads Function Equals(ByVal other As Combi) As Boolean Implements IEquatable(Of Combi).Equals
            'Handle same reference
            If (Object.ReferenceEquals(Me, other)) Then Return True
            'Handle null
            If (other Is Nothing) Then Return False
            'Compare count first
            If (Me.Count <> other.Count) Then Return False
            'Compare the flags
            Dim myFlagsX As Boolean() = Me.Flags
            Dim myFlagsY As Boolean() = other.Flags
            For i As Int32 = 0 To myFlagsX.Length - 1
                If (myFlagsX(i) <> myFlagsY(i)) Then Return False
            Next
            'Everything is the same, return true
            Return True
        End Function

        ''' <summary>Compares this combination with the given member and returns true if this combination consists only of
        ''' that member, false otherwise.</summary>
        ''' <param name="other">The member this combination is compared with</param>
        Public Overloads Function Equals(ByVal other As TEnum) As Boolean Implements IEquatable(Of TEnum).Equals
            If (other Is Nothing) Then Return False
            If (Count <> 1) Then Return False
            Dim myMember As TEnum = Item(0)
            Return (myMember.Equals(other))
        End Function

        ''' <summary>Retrieves an XOR combination of all set members. This ensures that a member and a Combination that contains 
        ''' only one member have the same hashcode.</summary>
        <EditorBrowsable(EditorBrowsableState.Advanced)> _
        Public Overrides Function GetHashCode() As Int32
            Dim myResult As Int32? = _HashCode
            If (myResult Is Nothing) Then
                'Calculate hashcode
                Dim myHashCode As Int32 = 0
                Dim myMembers As TEnum() = ToArray()
                For Each myMember As TEnum In myMembers
                    myHashCode = myHashCode Xor myMember.GetHashCode()
                Next
                'Assign and return
                _HashCode = myHashCode
                Return myHashCode
            End If
            'Return from cache
            Return myResult.Value
        End Function

        ''' <summary>Returns true if the given member is set, false otherwise (null also returns false).</summary>
        ''' <param name="member">The member to check</param>
        Public Function Contains(ByVal member As TEnum) As Boolean
            'Ignore if null
            If (member Is Nothing) Then Return False
            'Return whether the flag is set
            Return Flags(member.Index)
        End Function

        ''' <summary>Returns true if all given members are set, false otherwise (if the combination does not contain any members, 
        ''' false is returned).</summary>
        ''' <param name="members">The members to check</param>
        Public Function Contains(ByVal members As TCombi) As Boolean
            'Ignore if null
            If (members Is Nothing) OrElse (members.IsEmpty) Then Return False
            If (members.Count > Me.Count) Then Return False
            'Compare the flags
            Dim myFlags As Boolean() = Me.Flags
            Dim myMembers As TEnum() = members.ToArray()
            For Each myMember As TEnum In myMembers
                If (Not myFlags(myMember.Index)) Then Return False
            Next
            Return True
        End Function

        ''' <summary>Returns true if all given members are set, false otherwise (null values are ignored; if the array does not contain
        ''' any members, false is returned).</summary>
        ''' <param name="member">The members to check</param>
        Public Function Contains(ByVal ParamArray member As TEnum()) As Boolean
            'Ignore if null
            If (member Is Nothing) Then Return False
            'Set the according flags to false
            Dim myHasFlags As Boolean = False
            Dim myFlags As Boolean() = Me.Flags
            For Each myMember As TEnum In member
                If (myMember Is Nothing) Then Continue For
                If (Not myFlags(myMember.Index)) Then Return False
                myHasFlags = True
            Next
            'Return true if at least one flag was given
            Return myHasFlags
        End Function

        ''' <summary>Returns true if at least one of the given members is set, false otherwise. This operation is thread-safe. If 
        ''' members is null, false is returned.</summary>
        ''' <param name="members">The members to check</param>
        Public Function ContainsAny(ByVal members As TCombi) As Boolean
            'Ignore if null
            If (members Is Nothing) OrElse (members.IsEmpty) OrElse (Me.IsEmpty) Then Return False
            'Compare the flags
            Dim myFlags As Boolean() = Me.Flags
            Dim myMembers As TEnum() = members.ToArray()
            For Each myMember As TEnum In myMembers
                'Return true if set
                If (myFlags(myMember.Index)) Then Return True
            Next
            'Otherwise return false
            Return False
        End Function

        ''' <summary>Returns true if at least one of the given members is set, false otherwise. This operation is not thread-safe, 
        ''' please ensure the members collection is not manipulated during the time this method executes. If the collection is
        ''' null or empty, false is returned.</summary>
        ''' <param name="members">The members to check</param>
        Public Function ContainsAny(members As IEnumerable(Of TEnum)) As Boolean
            'Convert to array
            If (members Is Nothing) OrElse (Me.IsEmpty) Then Return False
            'Compare the flags
            Dim myFlags As Boolean() = Me.Flags
            For Each myMember As TEnum In members
                'Return true if set
                If (myMember Is Nothing) Then Continue For
                If (myFlags(myMember.Index)) Then Return True
            Next
            'Otherwise return false
            Return False
        End Function

        ''' <summary>Returns true if at least one of the given members is set, false otherwise. This operation is thread-safe. 
        ''' Null values are ignored.</summary>
        ''' <param name="member">The first member to check.</param>
        ''' <param name="additionalMembers">Additional members to check (split into two parameters for easier consumation by C#).</param>
        Public Function ContainsAny(member As TEnum, ByVal ParamArray additionalMembers As TEnum()) As Boolean
            'Handle empty array
            If (Me.IsEmpty) Then Return False
            Dim myFlags As Boolean() = Me.Flags
            If (additionalMembers Is Nothing) OrElse (additionalMembers.Length = 0) Then
                If (member Is Nothing) Then Return False
                Return myFlags(member.Index)
            End If
            'Compare the flags
            If (member IsNot Nothing) AndAlso (myFlags(member.Index)) Then Return True
            For Each myMember As TEnum In additionalMembers
                'Return true if set
                If (myMember Is Nothing) Then Continue For
                If (myFlags(myMember.Index)) Then Return True
            Next
            'Otherwise return false
            Return False
        End Function

        ''' <summary>Returns the set members as an array (.NET 2.0 support).</summary>
        Public Function ToArray() As TEnum()
            'Hint: May not be cached (or only for Array.Copy())
            Dim myResult As New List(Of TEnum)
            Dim myFlags As Boolean() = Flags
            Dim myMembers As TEnum() = Members
            For i As Int32 = 0 To myFlags.Length - 1
                If (myFlags(i)) Then myResult.Add(myMembers(i))
            Next
            Return myResult.ToArray()
        End Function

        Public Overrides Function ToString() As String
            'Write all set members
            Dim myResult As New StringBuilder()
            Dim myMembers As TEnum() = ToArray()
            If (myMembers.Length = 0) Then Return ""
            If (myMembers.Length = 1) Then Return myMembers(0).Name
            For Each myMember As TEnum In ToArray()
                myResult.Append(myMember.Name)
                myResult.Append(", ")
            Next
            myResult.Length -= 2
            Return myResult.ToString()
        End Function

        'Public Operators 

        ''' <summary>Every member is castable into a Combination. To avoid ambiguity for the operators this operation is
        ''' defined as explicite cast even if it cannot fail. This operation is thread-safe.</summary>
        ''' <param name="member">The member to be cast into a Combination.</param>
        Public Shared Narrowing Operator CType(ByVal member As TEnum) As Combi
            If (member Is Nothing) Then Return Combi.Empty
            Return member.ToCombi()
        End Operator

        ''' <summary>If the combination consists of exactly one member, it can be cast into the member. If no member or more than one
        ''' member is set, an InvalidCastException is thrown. This operation is thread-safe.</summary>
        ''' <param name="combination">The combination containing a single member to be cast.</param>
        ''' <exception cref="InvalidCastException">An InvalidCastException is thrown if the combination has less/more than one member set.</exception>
        Public Shared Narrowing Operator CType(ByVal combination As Combi) As TEnum
            'Check args
            If (combination Is Nothing) OrElse (combination.Count = 0) Then
                Throw New InvalidCastException("The combination has no members set.")
            End If
            If (combination.Count > 1) Then Throw New InvalidCastException("The combination has more than one member set.")
            'Return the result
            Dim myMembers As TEnum() = combination.ToArray()
            Return myMembers(0)
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. It is more efficient to use the construtor that takes a paramarray to combine more than two members.
        ''' This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a member.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Or(ByVal arg1 As Combi, ByVal arg2 As TEnum) As TCombi
            Return (arg2 Or arg1)
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a member.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Or(ByVal arg1 As TEnum, ByVal arg2 As Combi) As TCombi
            'Handle null
            If (arg1 Is Nothing) Then
                If (arg2 Is Nothing) Then Return Combi.Empty
                Return CType(arg2, TCombi)
            End If
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return arg1.ToCombi()
            'Return arg2 if the flag was already set
            If (arg2.Contains(arg1)) Then Return CType(arg2, TCombi)
            'Merge both flags
            Dim myResult As TCombi = Combi.GetInstanceOptimized(arg1, arg2.ToArray())
            Return myResult
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. If one of the arguments is null or empty, the other argument is returned (same instance). If both 
        ''' arguments are null, <see cref="Combi.Empty">Combi.Empty</see> is returned. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members (null is allowed).</param>
        ''' <param name="arg2">The second argument, a combination of members (null is allowed).</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Or(ByVal arg1 As Combi, ByVal arg2 As TCombi) As TCombi
            'Handle null
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) Then
                If (arg2 Is Nothing) Then Return Combi.Empty
                Return arg2
            End If
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return CType(arg1, TCombi) 'same instance
            'Handle equal
            If (arg1 = arg2) Then Return arg2
            'Handle contains
            Dim myCount1 As Int32 = arg1.Count
            Dim myCount2 As Int32 = arg2.Count
            If (myCount1 > myCount2) Then
                If (arg1.Contains(arg2)) Then
                    Return CType(arg1, TCombi)
                End If
            ElseIf (myCount1 < myCount2) Then
                If (arg2.Contains(CType(arg1, TCombi))) Then
                    Return arg2
                End If
            End If
            'Merge both flags
            Dim myResult As TCombi = Combi.GetInstanceRaw()
            Dim myFlags As Boolean() = myResult.Flags
            Dim myFlagsX As Boolean() = arg1.Flags
            Dim myFlagsY As Boolean() = arg2.Flags
            For i As Int32 = 0 To myFlags.Length - 1
                myFlags(i) = (myFlagsX(i) OrElse myFlagsY(i))
            Next
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination using a binary XOR operation. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a member.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Xor(ByVal arg1 As Combi, ByVal arg2 As TEnum) As TCombi
            'Handle null/empty
            If (arg2 Is Nothing) Then
                If (arg1 Is Nothing) Then Return Combi.Empty
                Return CType(arg1, TCombi)
            End If
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) Then Return arg2.ToCombi()
            'Merge both flags
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg1)
            myResult.Toggle(arg2)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination using a binary XOR operation. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a member.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Xor(ByVal arg1 As TEnum, ByVal arg2 As Combi) As TCombi
            'Handle null/empty
            If (arg1 Is Nothing) Then
                If (arg2 Is Nothing) Then Return Combi.Empty
                Return CType(arg2, TCombi)
            End If
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return arg1.ToCombi()
            'Merge both flags
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg2)
            myResult.Toggle(arg1)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination using a binary XOR operation. To toggle all
        ''' members use an unary NOT operation. If one of the arguments is null or empty, the other argument is returned 
        ''' (same instance). If both arguments are null, an empty combination is returned. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Xor(ByVal arg1 As Combi, ByVal arg2 As TCombi) As TCombi
            'Handle null/empty
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) Then
                If (arg2 Is Nothing) Then
                    Return Combi.Empty
                End If
                Return arg2 'same instance
            End If
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return CType(arg1, TCombi) 'same instance
            'Merge both flags
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg1)
            myResult.Toggle(arg2)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. It is more efficient to use the construtor that takes a paramarray to combine more than two members. 
        ''' This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a member.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator +(ByVal arg1 As Combi, ByVal arg2 As TEnum) As TCombi
            Return (arg1 Or arg2)
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a member.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator +(ByVal arg1 As TEnum, ByVal arg2 As Combi) As TCombi
            Return (arg1 Or arg2)
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator +(ByVal arg1 As Combi, ByVal arg2 As TCombi) As TCombi
            Return (arg1 Or arg2)
        End Operator

        ''' <summary>Returns a copy of arg1 (a new instance) without member arg2. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The argument to subtract, a member.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator -(ByVal arg1 As Combi, ByVal arg2 As TEnum) As TCombi
            'Handle null
            If (arg1 Is Nothing) Then Return Combi.Empty
            If (arg2 Is Nothing) Then Return CType(arg1, TCombi)
            'Return arg1 if the flag is not set
            If (Not arg1.Contains(arg2)) Then
                Return CType(arg1, TCombi)
            End If
            'Otherwise return new instance
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg1)
            myResult.Remove(arg2)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Returns a copy of arg1 (a new instance) without the members of arg2. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The argument to subtract, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator -(ByVal arg1 As Combi, ByVal arg2 As TCombi) As TCombi
            'Handle null
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) Then Return Combi.Empty
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return CType(arg1, TCombi)
            'Return same instance if there was no match
            If (Not arg1.ContainsAny(arg2)) Then
                Return CType(arg1, TCombi)
            End If
            'Remove the flags
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg1)
            myResult.Remove(arg2)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Returns a new instance of Combination that contains all members that are in both arg1 and arg2. This operation is
        ''' thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator And(ByVal arg1 As Combi, ByVal arg2 As TCombi) As TCombi
            'Handle null
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) OrElse (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return Combi.Empty
            'Handle no match
            If (Not arg1.ContainsAny(arg2)) Then
                Return Combi.Empty
            End If
            'Merge both flags
            Dim myResult As New TCombi()
            Dim myFlags As Boolean() = myResult.Flags
            Dim myFlagsX As Boolean() = arg1.Flags
            Dim myFlagsY As Boolean() = arg2.Flags
            For i As Int32 = 0 To myFlags.Length - 1
                myFlags(i) = (myFlagsX(i) AndAlso myFlagsY(i))
            Next
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Returns a new instance of Combination that contains all members that are not in arg. This 
        ''' operation is thread-safe.</summary>
        ''' <param name="arg">The argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Not(ByVal arg As Combi) As TCombi
            'Handle well-known collections
            If (arg Is Nothing) OrElse (arg.IsEmpty) Then Return Combi.All
            If (arg.IsAllSet) Then Return Combi.Empty
            'Return new inverted collection
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg)
            myResult.ToggleAll()
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Returns true if the given combination has only the given member set, false otherwise. This 
        ''' operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a combination of members.</param>
        ''' <param name="arg2">The second argument to compare, a member.</param>
        Public Shared Operator =(ByVal arg1 As Combi, ByVal arg2 As TEnum) As Boolean
            If (arg1 Is Nothing) Then Return False
            Return arg1.Equals(arg2)
        End Operator

        ''' <summary>Returns true if the given combination has only the given member set, false otherwise. This 
        ''' operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a member.</param>
        ''' <param name="arg2">The second argument to compare, a combination of members.</param>
        Public Shared Operator =(ByVal arg1 As TEnum, ByVal arg2 As Combi) As Boolean
            If (arg2 Is Nothing) Then Return False
            Return arg2.Equals(arg1)
        End Operator

        ''' <summary>Returns true if the given combinations have the same members. If both arguments are null
        ''' or empty (in any combination), true is returned. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a combination of members.</param>
        ''' <param name="arg2">The second argument to compare, a combination of members.</param>
        Public Shared Operator =(ByVal arg1 As Combi, ByVal arg2 As TCombi) As Boolean
            'Handle null
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) Then
                If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return True
                Return False
            End If
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return False
            'Compare the values
            Return arg1.Equals(arg2)
        End Operator

        ''' <summary>Returns false if the given combination has only the given member set, true otherwise. (If the member is null
        ''' and also the combination is null or empty, false is returned to be consistent to the explicite cast operator). This 
        ''' operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a combination of members.</param>
        ''' <param name="arg2">The second argument to compare, a member.</param>
        Public Shared Operator <>(ByVal arg1 As Combi, ByVal arg2 As TEnum) As Boolean
            Return (Not (arg1 = arg2))
        End Operator

        ''' <summary>Returns false if the given combination has only the given member set, true otherwise. (If the member is null
        ''' and also the combination is null or empty, false is returned to be consistent to the explicite cast operator). This 
        ''' operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a member.</param>
        ''' <param name="arg2">The second argument to compare, a combination of members.</param>
        Public Shared Operator <>(ByVal arg1 As TEnum, ByVal arg2 As Combi) As Boolean
            Return (Not (arg1 = arg2))
        End Operator

        ''' <summary>Returns false if the given combinations have the same members. If both combinations are null or empty (in any 
        ''' combination), false is returned. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a combination of members.</param>
        ''' <param name="arg2">The second argument to compare, a combination of members.</param>
        Public Shared Operator <>(ByVal arg1 As Combi, ByVal arg2 As TCombi) As Boolean
            Return (Not (arg1 = arg2))
        End Operator

        'Private Properties

        ''' <summary>Initializes and returns the flags.</summary>
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private ReadOnly Property Flags As Boolean()
            Get
                Dim myResult As Boolean() = _Flags
                If (myResult Is Nothing) Then
                    myResult = New Boolean(CustomEnum(Of TEnum, TCombi).Members.Length - 1) {}
                    _Flags = myResult
                End If
                Return myResult
            End Get
        End Property

        ''' <summary>For debugging only, visualization for property Flags.</summary>
        Private ReadOnly Property AllMembers() As MemberFlag()
            Get
                Dim myMembers As TEnum() = GetMembers()
                Dim myFlags As Boolean() = Flags
                Dim myResult(myMembers.Length - 1) As MemberFlag
                For i As Int32 = 0 To myMembers.Length - 1
                    myResult(i) = New MemberFlag(myMembers(i), myFlags(i))
                Next
                Return myResult
            End Get
        End Property

        ''' <summary>For debugging only, calls function ToArray().</summary>
        Private ReadOnly Property AssignedMembers() As TEnum()
            Get
                Return ToArray()
            End Get
        End Property

        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private ReadOnly Property DebuggerDisplayValue As String
            Get
                'Write all set members
                Dim myResult As New StringBuilder()
                Dim myMembers As TEnum() = ToArray()
                If (myMembers.Length = 0) Then Return "[empty]"
                If (myMembers.Length = 1) Then Return myMembers(0).Name
                For Each myMember As TEnum In ToArray()
                    myResult.Append(myMember.Name)
                    myResult.Append(", ")
                Next
                myResult.Length -= 2
                Return myResult.ToString()
            End Get
        End Property

        'Private Methods

        ''' <summary>Why optimize the result? Because 1) comparing same instances is much faster than comparing instances
        ''' that are only equal, 2) garbage collection of level 0 objects is rather cheap, 3) we use a little less memory
        ''' and 4) it is does not cost much to perform these optimizations (the initialization of Count is the only somewhat
        ''' costy operation).</summary>
        ''' <param name="result">A combination containing the same members.</param>
        Private Shared Function OptimizeResult(ByVal result As TCombi) As TCombi
            If (result Is Nothing) Then Return Combi.Empty
            Select Case result.Count
                Case 0
                    Return Combi.Empty
                Case 1
                    Return result(0).ToCombi()
                Case Else
                    If (result.IsAllSet) Then Return Combi.All
            End Select
            Return result
        End Function

        ''' <summary>Sets the given member (if the member is already set or <paramref name="member" /> is null, it is ignored). 
        ''' An <see cref="InvalidOperationException" /> is thrown if the instance is protected. The Set method is similair to a 
        ''' binary OR operation.</summary>
        ''' <param name="member">The member to set.</param>
        Private Sub [Set](ByVal member As TEnum)
            'Ignore if null
            If (member Is Nothing) Then Return
            'Set the flag
            Flags(member.Index) = True
        End Sub

        ''' <summary>Sets the given members (null values and duplicates are ignored). An <see cref="InvalidOperationException" /> 
        ''' is thrown if the instance is protected. The Set method is similair to a binary OR operation.</summary>
        ''' <param name="member">The members to set.</param>
        Private Sub [Set](ByVal member As TEnum())
            'Ignore if null
            If (member Is Nothing) Then Return
            'Set the according flags
            Dim myFlags As Boolean() = Me.Flags
            For Each myMember As TEnum In member
                If (myMember Is Nothing) Then Continue For
                myFlags(myMember.Index) = True
            Next
        End Sub

        ''' <summary>Sets all members. An <see cref="InvalidOperationException" /> is thrown if the instance is protected.</summary>
        Private Sub SetAll()
            'Set all flags to true
            Dim myFlags As Boolean() = Flags
            For i As Int32 = 0 To myFlags.Length - 1
                myFlags(i) = True
            Next
        End Sub

        ''' <summary>Toggles a member (if it is set, it is removed; if it is not set, it is set). If the member is null, it is ignored. 
        ''' An <see cref="InvalidOperationException" /> is thrown if the instance is protected. Toggle is similair to a binary 
        ''' XOR operation.</summary>
        ''' <param name="member">The member to toggle.</param>
        Private Sub Toggle(ByVal member As TEnum)
            'Ignore if null
            If (member Is Nothing) Then Return
            'Toggle the flag
            Dim myFlags As Boolean() = Flags
            Dim myIndex As Int32 = member.Index
            myFlags(myIndex) = (Not myFlags(myIndex))
        End Sub

        ''' <summary>Merges the given combination using a binary XOR operation.</summary>
        ''' <param name="combination">The members to toggle.</param>
        Private Sub Toggle(ByVal combination As Combi)
            'Handle null
            If (combination Is Nothing) OrElse (combination.IsEmpty) Then Return
            'Merge both flags
            Dim myFlagsX As Boolean() = Flags
            Dim myFlagsY As Boolean() = combination.Flags
            For i As Int32 = 0 To myFlagsX.Length - 1
                Dim myX As Boolean = myFlagsX(i)
                Dim myY As Boolean = myFlagsY(i)
                If (myX AndAlso myY) Then
                    myFlagsX(i) = False
                ElseIf (myX OrElse myY) Then
                    myFlagsX(i) = True
                End If
            Next
        End Sub

        ''' <summary>Toggles all members. An <see cref="InvalidOperationException" /> is thrown if the instance is protected. ToggleAll is 
        ''' similair to an unary NOT operation.</summary>
        Private Sub ToggleAll()
            'Toggle all flags
            Dim myFlags As Boolean() = Flags
            For i As Int32 = 0 To myFlags.Length - 1
                myFlags(i) = (Not myFlags(i))
            Next
        End Sub

        ''' <summary>Removes the given member (if the member is not set or <paramref name="member" /> is null, it is ignored).
        ''' An <see cref="InvalidOperationException" /> is thrown if the instance is protected. The Remove method is similair to 
        ''' a binary NOT operation.</summary>
        ''' <param name="member">The member to remove</param>
        Private Sub Remove(ByVal member As TEnum)
            'Ignore if null
            If (member Is Nothing) Then Return
            'Remove the flag
            Flags(member.Index) = False
        End Sub

        ''' <summary>Removes the given members (null values and duplicates are ignored). An <see cref="InvalidOperationException" /> is 
        ''' thrown if the instance is protected. The Remove method is similair to a binary NOT operation (does not exist in Visual Basic).</summary>
        ''' <param name="combination">The members to remove</param>
        Private Sub Remove(ByVal combination As Combi)
            'Ignore if null
            If (combination Is Nothing) OrElse (combination.IsEmpty) Then Return
            'Set the according flags to false
            Dim myFlags As Boolean() = Me.Flags
            Dim myOtherMembers As TEnum() = combination.ToArray()
            For Each myMember As TEnum In myOtherMembers
                myFlags(myMember.Index) = False
            Next
        End Sub

        ''' <summary>Retrieves an enumerator for the set members. This method takes a snapshot of the members before returning the 
        ''' enumerator, allowing you to manipulate the members even from within the loop.</summary>
        Private Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of TEnum) Implements System.Collections.Generic.IEnumerable(Of TEnum).GetEnumerator
            Dim myResult As IEnumerable(Of TEnum) = ToArray()
            Return myResult.GetEnumerator()
        End Function

        ''' <summary>Retrieves an enumerator for the set members. This method takes a snapshot of the members before returning the 
        ''' enumerator, allowing you to manipulate the members even from within the loop.</summary>
        Private Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function


        '**********************************************************************
        ' INNER CLASS: MemberFlag
        '**********************************************************************

        ''' <summary>For debugged display.</summary>
        <DebuggerDisplay("{Flag}", Name:="{Name}")> _
        Private Class MemberFlag

            'Public Fields

            <SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields", Justification:="It is used to be displayed in the debugger.")> _
            <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
            Public ReadOnly Name As String
            <SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields", Justification:="It is used to be displayed in the debugger.")> _
            <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
            Public ReadOnly Flag As Boolean
            <SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields", Justification:="It is used to be displayed in the debugger.")> _
            Public ReadOnly Member As TEnum

            'Constructors

            Public Sub New(ByVal member As TEnum, ByVal flag As Boolean)
                Me.Name = member.Name
                Me.Flag = flag
                Me.Member = member
            End Sub

        End Class

    End Class

End Class


''' <summary>Base class of valueless custom enumerations that do not need to overwrite the combination class.</summary>
<SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix", Justification:="Because it is a base class for all enums, the suffix ""Enum"" is just fine.")> _
Public MustInherit Class CustomEnum(Of TEnum As CustomEnum(Of TEnum))
    Inherits CustomEnum(Of TEnum, Combi)

    'Constructors

    Protected Sub New()
        MyBase.New()
    End Sub

    Protected Sub New(ByVal caseSensitive As Boolean?)
        MyBase.New(caseSensitive)
    End Sub


    '**********************************************************************
    ' INNER CLASS: Combi
    '**********************************************************************

    <SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix", Justification:="That would make it even longer!")> _
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1034:NestedTypesShouldNotBeVisible")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shadows Class Combi
        Inherits CustomEnum(Of TEnum, Combi).Combi

    End Class

End Class



''' <summary>
''' Base-class for generic custom enums that take a value and need to overwrite the combination class.
''' </summary>
''' <typeparam name="TEnum">The type of your subclass</typeparam><remarks>
''' Hints for implementors: You must ensure that only one instance of each enum-value exists. This is easily reached by
''' declaring the constructor(s) private, sealing the class and exposing the enum-values as static fields. If you are
''' implementing them through static getter properties, make sure lazy initialization is used and that always the same
''' instance is returned with every call.
''' </remarks>
<SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
<SuppressMessage("Microsoft.Design", "CA1005:AvoidExcessiveParametersOnGenericTypes", Justification:="It is needed, period.")> _
<SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix", Justification:="Because it is a base class for all enums, the suffix ""Enum"" is just fine.")> _
<SuppressMessage("Microsoft.Design", "CA1046:DoNotOverloadOperatorEqualsOnReferenceTypes", Justification:="CustomEnum behaves like a value type.")> _
<DebuggerDisplay("{DebuggerDisplayValue}")> _
Public MustInherit Class ValueEnum(Of TEnum As {ValueEnum(Of TEnum, TValue, TCombi)}, TValue, TCombi As {ValueEnum(Of TEnum, TValue, TCombi).Combi, New})
    Inherits CustomEnum
    Implements IEquatable(Of TValue)

    'Private Fields

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared _Members As TEnum()
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared _CaseSensitive As Boolean?
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared _NameComparer As IEqualityComparer(Of String)
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared _ValueComparer As IEqualityComparer(Of TValue)
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared _IsFirstInstance As Boolean = True
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared _Lock As New Object()
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private _Name As String
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private _Value As TValue
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private _Index As Int32 = -1
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private _Combination As TCombi = Nothing

    'Constructors

    ''' <summary>Called by implementors to create a new instance of TEnum (when assigning the instance to a static field). 
    ''' Important: Make your constructors private to ensure there are no instances except the ones initialized 
    ''' by your subclass! Null values are not supported and throw an ArgumentNullException.</summary>
    ''' <param name="value">This member's value (not null)</param>
    ''' <exception cref="ArgumentNullException">An ArgumentNullException is thrown if aValue is null.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this instance's type is not of type <typeparam name="TEnum" />.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException may be thrown (check only in full-trust possible) if this instance has non-private constructors.</exception>
    Protected Sub New(ByVal value As TValue)
        Me.New(value, Nothing, Nothing)
    End Sub

    ''' <summary>Called by implementors to create a new instance of TEnum (when assigning the instance to a static field). 
    ''' Important: Make your constructors private to ensure there are no instances except the ones initialized 
    ''' by your subclass! Null values are not supported and throw an ArgumentNullException.</summary>
    ''' <param name="value">This member's value (not null)</param>
    ''' <param name="caseSensitive">Leave null for automatic determination (recommended), or set explicitely.</param>
    ''' <param name="valueComparer">What comparer should be used to compare the values in method <see cref="GetMemberByValue" /> and <see cref="Equals" /> as well as the equal operator or null to use Object.Equals(..).</param>
    ''' <exception cref="ArgumentNullException">An ArgumentNullException is thrown if aValue is null.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this instance's type is not of type <typeparam name="TEnum" />.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException may be thrown (check only in full-trust possible) if this instance has non-private constructors.</exception>
    Protected Sub New(ByVal value As TValue, ByVal caseSensitive As Boolean?, ByVal valueComparer As IEqualityComparer(Of TValue))
        'Check the value
        If (value Is Nothing) Then Throw New ArgumentNullException("value")
        'Assign instance value
        _Value = value
        'Make sure no evil cross subclass is changing our static variable
        If (Not GetType(TEnum).IsAssignableFrom(Me.GetType())) Then
            Throw New InvalidOperationException("Internal error in " & Me.GetType().Name & "! Change the first type parameter from """ & GetType(TEnum).Name & """ to """ & Me.GetType().Name & """.")
        End If
        'Make sure only the first subclass is affecting our static variables
        If (_IsFirstInstance) Then
            _IsFirstInstance = False
            'Check constructors
            CheckConstructors()
            'Assign static variables
            _CaseSensitive = caseSensitive
            _ValueComparer = valueComparer
        End If
    End Sub

    'Public Properties

    ''' <summary>Returns the name of this member (the name of the static field or static getter property this member 
    ''' was assigned to in the subclass). Watch out: Do not access from within the subclasses constructor!</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this property is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is also thrown if this instance is not 
    ''' assigned to a public static readonly field or property of TEnum.</exception>
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="readonly")> _
    Public Overrides ReadOnly Property Name As String
        Get
            'Check whether the name is already assigned
            Dim myResult As String = _Name
            If (myResult Is Nothing) Then
                InitAndReturnMembers()
                myResult = _Name
                If (myResult Is Nothing) Then
                    Throw New InvalidOperationException("Detached instance error! Ensure that all members are assigned to a public static readonly field or property.")
                End If
            End If
            'Return the name
            Return myResult
        End Get
    End Property

    ''' <summary>Returns the value of this member as specified in the constructor.</summary>
    Public ReadOnly Property Value As TValue
        Get
            Return _Value
        End Get
    End Property

    ''' <summary>Returns the zero-based index position of this member. Watch out: Do not access from within the subclasses constructor!
    ''' If there are static fields as well as static getter properties, the fields have the lower index. The order is the same as it 
    ''' is returnd from Type.GetFields() and Type.GetProperties() and should correspond to the order the fields/properties have been 
    ''' declared. The index may be used by to compare members.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this property is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is also thrown if this instance is not 
    ''' assigned to a public static readonly field or property of TEnum.</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="readonly")> _
    Public Overrides ReadOnly Property Index As Int32
        Get
            'Check whether the index is already assigned
            Dim myResult As Int32 = _Index
            If (myResult < 0) Then
                InitAndReturnMembers()
                myResult = _Index
                If (myResult < 0) Then
                    Throw New InvalidOperationException("Detached instance error! Ensure that all members are assigned to a public static readonly field or property.")
                End If
            End If
            'Return the index
            Return myResult
        End Get
    End Property

    'Public Methods

    ''' <summary>Returns true if one of the given members equals this member, false otherwise. If the given paramarray 
    ''' is null, false is returned. If the array contains null values they are ignored.</summary>
    Public Function [In](ByVal ParamArray members As TEnum()) As Boolean
        'Check the args
        If (members Is Nothing) OrElse (members.Length = 0) Then Return False
        'Loop through given members
        Dim myInstance As TEnum = DirectCast(Me, TEnum)
        For Each myMember As TEnum In members
            If (myMember Is Nothing) Then Continue For
            If (myMember.Equals(myInstance)) Then Return True
        Next
        'Otherwise return false
        Return False
    End Function

    ''' <summary>Returns true if the given combination contains this member, false otherwise. False is also returned if
    ''' the combination is null.</summary>
    ''' <param name="combination">The combination to check (null is allowed).</param>
    Public Function [In](ByVal combination As TCombi) As Boolean
        'Check args
        If (combination Is Nothing) Then Return False
        'Determine whether it contains us
        Dim myInstance As TEnum = DirectCast(Me, TEnum)
        Return combination.Contains(myInstance)
    End Function

    ''' <summary>Returns the member with the next higher index. Watch out: Do not access from within the subclasses constructor!
    ''' If this member is the last one, then it depends on parameter <paramref name="loop" /> whether the first member is returned 
    ''' (<paramref name="loop" /> is set to <c>true</c>) or null (<paramref name="loop" /> is set to <c>false</c>). If the enum does 
    ''' not contain any members, null is returned.</summary>
    ''' <param name="loop">Whether the first member should be returned if the end is reached (true) or null (false).</param>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this method is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this method is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Function GetNext(ByVal [loop] As Boolean) As TEnum
        Dim myIndex As Int32 = Index + 1
        Dim myMembers As TEnum() = Members
        If (myIndex >= myMembers.Length) Then
            If ([loop]) Then Return myMembers(0)
            Return Nothing
        End If
        Return myMembers(myIndex)
    End Function

    ''' <summary>Returns the member with the next lower index. If this member is the first one, then it depends on parameter 
    ''' <paramref name="loop" /> whether the last member is returned (<paramref name="loop" /> is set to <c>true</c>) or null 
    ''' (<paramref name="loop" /> is set to <c>false</c>). If the enum does not contain any members, null is returned.</summary>
    ''' <param name="loop">Whether the last member should be returned if the start is reached (true) or null (false).</param>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Function GetPrevious(ByVal [loop] As Boolean) As TEnum
        Dim myIndex As Int32 = Index - 1
        Dim myMembers As TEnum() = Members
        If (myIndex < 0) Then
            If ([loop]) Then Return myMembers(myMembers.Length - 1)
            Return Nothing
        End If
        Return myMembers(myIndex)
    End Function

    'Public Class Properties

    ''' <summary>Returns the type of the combination class.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared ReadOnly Property CombiType As Type
        Get
            Return GetType(TCombi)
        End Get
    End Property

    ''' <summary>Returns whether the names passed to function <see cref="GetMemberByName" /> are treated case sensitive
    ''' or not (using <see cref="StringComparer.Ordinal" /> resp. <see cref="StringComparer.OrdinalIgnoreCase" />).
    ''' The default behavior is that they are case-insensitive except there would be two or more entries that would
    ''' cause an ambiguity.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException may be thrown if this property is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this property is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared ReadOnly Property CaseSensitive As Boolean
        Get
            Dim myResult As Boolean? = _CaseSensitive
            If (myResult Is Nothing) Then
                Return InitAndReturnCaseSensitive(Members)
            End If
            Return myResult.Value
        End Get
    End Property

    ''' <summary>Returns the underlying type (the type of the <see cref="Value" /> property).</summary>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared ReadOnly Property ValueType As Type
        Get
            Return GetType(TValue)
        End Get
    End Property

    ''' <summary>Gets the first defined member of this enum (the one with index 0). If the enum does not contain any members,
    ''' null is returned.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this property is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this property is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Public Shared ReadOnly Property First As TEnum
        Get
            Dim myMembers As TEnum() = Members
            If (myMembers.Length = 0) Then Return Nothing
            Return myMembers(0)
        End Get
    End Property

    ''' <summary>Gets the last defined member of this enum (the one with the highest index). If the enum does not contain any members,
    ''' null is returned.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this property is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this property is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Public Shared ReadOnly Property Last As TEnum
        Get
            Dim myMembers As TEnum() = Members
            If (myMembers.Length = 0) Then Return Nothing
            Return myMembers(myMembers.Length - 1)
        End Get
    End Property

    'Public Class Functions

    ''' <summary>Returns an empty combination (a reference to <see cref="Combi.Empty">Combi.Empty</see>).-</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi() As TCombi
        Return Combi.Empty
    End Function

    ''' <summary>Returns a combination containing the member (if the member is null, a reference to 
    ''' <see cref="Combi.Empty">Combi.Empty</see> is returned, otherwise <see cref="ToCombi">member.ToCombi()</see>).</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi(ByVal member As TEnum) As TCombi
        Return Combi.GetInstanceOptimized(member)
    End Function

    ''' <summary>Returns a combination containing the given members.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi(ByVal member1 As TEnum, ByVal ParamArray member2 As TEnum()) As TCombi
        Return Combi.GetInstanceOptimized(member1, member2)
    End Function

    ''' <summary>Returns and instance that contains the given members. If members is null or empty, 
    ''' <see cref="Combi.Empty">Combi.Empty</see> is returned, if the member contains exactly one non-null-member,
    ''' <see cref="ToCombi">members[0].ToCombi()</see> is returned, otherwise a new instance containing 
    ''' the given members.</summary>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi(ByVal members As TEnum()) As TCombi
        Return Combi.GetInstanceOptimized(members)
    End Function

    ''' <summary>Returns and instance that contains the given members. If members is null or empty, 
    ''' <see cref="Combi.Empty">Combi.Empty</see> is returned, otherwise a reference to the given member.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi(ByVal members As Combi) As TCombi
        Return Combi.GetInstanceOptimized(members)
    End Function

    ''' <summary>Constructs a new combination and initializes the provided members. Watch out: This function is
    ''' not thread-safe, use synchronization means to prevent other threads from manipulating the collection
    ''' until the combination is returned.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetCombi(ByVal memberCollection As IEnumerable(Of TEnum)) As TCombi
        If (memberCollection Is Nothing) Then Return Combi.Empty
        Dim myArray As TEnum() = ToArray(Of TEnum)(memberCollection)
        Return Combi.GetInstanceOptimized(myArray)
    End Function

    ''' <summary>Returns a new array containing all members defined by this enum (in index order).</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMembers() As TEnum()
        Dim myMembers As TEnum() = Members
        Dim myResult(myMembers.Length - 1) As TEnum
        Array.Copy(myMembers, myResult, myMembers.Length)
        Return myResult
    End Function

    ''' <summary>Returns the names of all members defined by this enum (in index order). The names are the ones of
    ''' the "public static fields" and "public static getter properties without indexer" of this CustomEnum's type.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetNames() As String()
        Dim myMembers As TEnum() = Members
        Dim myResult(myMembers.Length - 1) As String
        For i As Int32 = 0 To myMembers.Length - 1
            myResult(i) = myMembers(i)._Name 'speed optimized (property is always initialized)
        Next
        Return myResult
    End Function

    ''' <summary>Returns the values of all enum values defined by this enum (in index order).</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetValues() As TValue()
        Dim myMembers As TEnum() = Members
        Dim myResult(myMembers.Length - 1) As TValue
        For i As Int32 = 0 To myMembers.Length - 1
            myResult(i) = myMembers(i)._Value
        Next
        Return myResult
    End Function

    ''' <summary>
    ''' Returns the member of the given name or null if not found. Property <see cref="CaseSensitive" /> tells whether 
    ''' <paramref name="name" /> is treated case-sensitive or not. If name is null, an ArgumentNullException is thrown. 
    ''' If the subclass is incorrectly implemented and has duplicate names defined, an InvalidOperationException is thrown. 
    ''' This function is thread-safe.
    ''' </summary>
    ''' <param name="name">The name to look up.</param>
    ''' <returns>The member with the given name of null if not found.</returns>
    ''' <exception cref="ArgumentNullException">An ArgumentNullException is thrown if <paramref name="name" /> is null.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if the subclass explicitely 
    ''' defined the enum to be case-insensitive but the names are ambiguous.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    ''' <remarks>A full duplicate check is performed the first time this method is called.</remarks>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMemberByName(ByVal name As String) As TEnum
        'Check the argument
        If (name Is Nothing) Then Throw New ArgumentNullException("name")
        'Get/initialize members and comparer
        Dim myMembers As TEnum() = Members
        Dim myComparer As IEqualityComparer(Of String) = NameComparer
        'Return the first name found (it is always unique, ensured when the NameComparer is initialized)
        For Each myMember As TEnum In myMembers
            If (myComparer.Equals(myMember._Name, name)) Then Return myMember 'speed optimized, _Name is always initialized
        Next
        'Otherwise return null
        Return Nothing
    End Function

    ''' <summary>
    ''' Returns the member of the given name or null if not found using the given comparer
    ''' to perform the comparison. An ArgumentException is thrown if the result would be ambiguous
    ''' according to the given comparer. If there are no special reasons don't use this method but the
    ''' one without the comparer overload as it is optimized to perform the duplicate check only
    ''' once and not every time the method is used. This method is thread-safe if the nameComparer
    ''' is thread-safe (or not being manipulated during the method call).</summary>
    ''' <param name="name">The name to look up.</param>
    ''' <param name="nameComparer">The comparer to use for the equality comparison of the strings (null defaults to StringComparer.Ordinal).</param>
    ''' <returns>The member or null if not found (or throws an exception if more than one is found).</returns>
    ''' <exception cref="ArgumentNullException">An ArgumentNullException is thrown if name is null.</exception>
    ''' <exception cref="ArgumentException">An ArgumentException is thrown if the result would be ambiguous according to the given comparer.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMemberByName(ByVal name As String, ByVal nameComparer As IEqualityComparer(Of String)) As TEnum
        'Check the argument
        If (name Is Nothing) Then Throw New ArgumentNullException("name")
        'Use optimized method if possible
        If (nameComparer Is Nothing) Then nameComparer = StringComparer.Ordinal
        If (Object.Equals(nameComparer, _NameComparer)) Then Return GetMemberByName(name)
        'Get the members
        Dim myMembers As TEnum() = Members
        'Get the first found member but continue looping
        Dim myResult As TEnum = Nothing
        For Each myMember As TEnum In myMembers
            If (nameComparer.Equals(myMember._Name, name)) Then
                If (myResult Is Nothing) Then
                    myResult = myMember
                Else
                    Throw New ArgumentException("According to the given comparer at least two ambiguous matches were found!")
                End If
            End If
        Next
        'Return the result 
        Return myResult
    End Function

    ''' <summary>
    ''' Returns the found members with the given names or Combi.Empty if not found. Property <see cref="CaseSensitive" /> 
    ''' tells whether the parameters are treated case-sensitive or not. For easier consumation by other languages two parameters
    ''' are used instead of one. If both arguments are null, empty or contain only unknown members, Combi.Empty is returned. 
    ''' Duplicates and null values within the array are ignored. If the subclass is incorrectly implemented and has duplicate names 
    ''' defined, an InvalidOperationException is thrown. This function is thread-safe.
    ''' </summary>
    ''' <param name="name">The name to look up.</param>
    ''' <param name="additionalNames">The name to look up.</param>
    ''' <returns>The found members as a combination.</returns>
    ''' <remarks>A full duplicate check is performed the first time this method (or GetMemberByName) is called.</remarks>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId:="ByNames", Justification:="The spelling is okay like this.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMembersByNames(ByVal name As String, ByVal ParamArray additionalNames As String()) As TCombi
        'Handle single value
        If (additionalNames Is Nothing) OrElse (additionalNames.Length = 0) Then
            'Handle null/empty
            If (name Is Nothing) Then Return Combi.Empty
            'Get member
            Dim myMember As TEnum = GetMemberByName(name)
            If (myMember Is Nothing) Then Return Combi.Empty
            Return myMember.ToCombi()
        End If
        'Get/initialize comparer and dictionary
        Dim myComparer As IEqualityComparer(Of String) = NameComparer
        Dim myDict As New Dictionary(Of String, String)(additionalNames.Length + 1, myComparer)
        'Fill in names
        If (name IsNot Nothing) Then myDict.Add(name, name)
        For Each myName As String In additionalNames
            If (myName Is Nothing) Then Continue For
            myDict.Item(myName) = myName
        Next
        'Return the result
        Return GetMembersByNames(myDict)
    End Function

    ''' <summary>
    ''' Returns the found members with the given names or Combi.Empty if not found. Property <see cref="CaseSensitive" /> 
    ''' tells whether the names are treated case-sensitive or not. If the enumeration is null, empty or contains only unknown 
    ''' members, Combi.Empty is returned. Duplicates and null values within the collection are ignored. If the subclass is 
    ''' incorrectly implemented and has duplicate names defined, an InvalidOperationException is thrown. This function is 
    ''' thread-safe only if the given nameCollection is thread-safe (or not manipulated during the time this function executes).
    ''' </summary>
    ''' <param name="nameCollection">The names to look up.</param>
    ''' <returns>The found members as a combination.</returns>
    ''' <remarks>A full duplicate check is performed the first time this method (or GetMemberByName) is called.</remarks>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId:="ByNames", Justification:="The spelling is okay like this.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMembersByNames(ByVal nameCollection As IEnumerable(Of String)) As TCombi
        'Handle null
        If (nameCollection Is Nothing) Then Return Combi.Empty
        'Get/initialize comparer and dictionary
        Dim myComparer As IEqualityComparer(Of String) = NameComparer
        Dim myDict As Dictionary(Of String, String) = Nothing
        Dim myCollection As ICollection(Of String) = TryCast(nameCollection, ICollection(Of String))
        If (myCollection Is Nothing) Then
            myDict = New Dictionary(Of String, String)(myComparer)
        Else
            myDict = New Dictionary(Of String, String)(myCollection.Count, myComparer)
        End If
        'Fill in names
        For Each myName As String In nameCollection
            If (myName Is Nothing) Then Continue For
            myDict.Item(myName) = myName
        Next
        'Return the result
        Return GetMembersByNames(myDict)
    End Function

    ''' <summary>
    ''' Returns the members of the given names. If the a name is not found or null, it is ignored. If there are ambiguous matches according to the
    ''' given comparer, an ArgumentException is thrown. If nameComparer is null, the comparison is performed using an ordinal comparer. If there are 
    ''' no special reasons don't use this method but the one without the comparer overload as it is optimized to perform the duplicate check only
    ''' once and not every time the method is called. This method is thread-safe if the nameComparer is thread-safe (or not being manipulated during 
    ''' the method call).</summary>
    ''' <param name="nameCollection">The names to look up.</param>
    ''' <param name="nameComparer">The comparer to use for the equality comparison of the strings (null defaults to StringComparer.Ordinal).</param>
    ''' <returns>The members found as a combination (an ArgumentException is thrown if there are duplicates).</returns>
    ''' <exception cref="ArgumentException">An ArgumentException is thrown if the result would be ambiguous according to the given comparer.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId:="ByNames", Justification:="The spelling is okay like this.")> _
    <SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId:="0", Justification:="Null values are allowed.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMembersByNames(ByVal nameCollection As IEnumerable(Of String), ByVal nameComparer As IEqualityComparer(Of String)) As TCombi
        'Handle null
        If (nameCollection Is Nothing) Then Return Combi.Empty
        'Get/initialize comparer and dictionary
        If (nameComparer Is Nothing) Then nameComparer = StringComparer.Ordinal
        Dim myDict As Dictionary(Of String, String) = Nothing
        Dim myCollection As ICollection(Of String) = TryCast(nameCollection, ICollection(Of String))
        If (myCollection Is Nothing) Then
            myDict = New Dictionary(Of String, String)(nameComparer)
        Else
            myDict = New Dictionary(Of String, String)(myCollection.Count, nameComparer)
        End If
        'Fill in names
        For Each myName As String In nameCollection
            If (myName Is Nothing) Then Continue For
            myDict.Item(myName) = myName
        Next
        'Return the result
        Return GetMembersByNames(myDict)
    End Function

    ''' <summary>
    ''' Gets the member by index. If the index is out-of-bounds null is returned. This function is faster
    ''' that calling GetMembers() with the index because function GetMembers has to copy the array and this
    ''' function does not need to.
    ''' </summary>
    ''' <param name="index">The zero-based index of the member.</param>
    ''' <returns>The member at the given zero-based index (null if out-of-range)</returns>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMemberByIndex(ByVal index As Int32) As TEnum
        If (index < 0) Then Return Nothing
        Dim myMembers As TEnum() = Members
        If (index >= myMembers.Length) Then Return Nothing
        Return myMembers(index)
    End Function

    ''' <summary>
    ''' Gets the members by the given indices. If some of the indices are out-of-bounds, they are ignored. If the array is null or
    ''' empty, it is ignored. This function is faster that calling GetMembers() with the index because function GetMembers has to 
    ''' copy the array and this function does not need to.
    ''' </summary>
    ''' <param name="index">The zero-based index of the first member to look up.</param>
    ''' <param name="additionalIndexes">Additional indices to look up (split into two parameters for better consumation by C#).</param>
    ''' <returns>A combination with the members of the given indices set.</returns>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMembersByIndexes(ByVal index As Int32, ByVal ParamArray additionalIndexes As Int32()) As TCombi
        'Return the result
        Return Combi.GetInstanceByIndexesOptimized(index, additionalIndexes)
    End Function

    ''' <summary>
    ''' Returns the found members with the given indices. If the enumerable is null, empty or contains only indices that are
    ''' out-of-range, Combi.Empty is returned. Duplicate values within the collection are ignored. If the subclass is 
    ''' incorrectly implemented and has duplicate names defined, an InvalidOperationException is thrown. This function is 
    ''' thread-safe only if the given indexCollection is thread-safe (or not manipulated during the time this function executes).
    ''' </summary>
    ''' <param name="indexCollection">The indices to look up.</param>
    ''' <returns>The found members as a combination.</returns>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMembersByIndexes(ByVal indexCollection As IEnumerable(Of Int32)) As TCombi
        Dim myArray As Int32() = ToArray(Of Int32)(indexCollection)
        Return Combi.GetInstanceByIndexesOptimized(myArray)
    End Function

    ''' <summary>Returns the first member of the given value or null if not found. If the given value is null, an ArgumentNullException
    ''' is thrown. This function uses the value comparer defined by the enumeration (or Object's  equal method if not defined) to 
    ''' perform the comparison. This method is thead-safe in a way that if comparing the values throws an exception it is swallowed and 
    ''' ignored. If this is not suitable for you make sure the value and the value-comparer that was possibly provided in the 
    ''' subclasses constructor are thread-safe (or not being manipulated during the function call).</summary>
    ''' <param name="value">The value to look up (not null).</param>
    ''' <returns>The enum entry or null if not found.</returns>
    ''' <exception cref="ArgumentNullException">An ArgumentNullException is thrown if <paramref name="value" /> is null.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMemberByValue(ByVal value As TValue) As TEnum
        Return GetMemberByValue(value, ValueComparer)
    End Function

    ''' <summary>Returns the first member of the given value or null if not found. If the given value is null, an ArgumentNullException 
    ''' is thrown. For comparing the values, the provided comparer is used. This method is thead-safe in a way that if comparing 
    ''' the values throws an exception it is swallowed and ignored. If this is not suitable for you make sure the value and the 
    ''' value-comparer that was possibly provided in the subclasses constructor are thread-safe (or not being manipulated during 
    ''' the function call).</summary>
    ''' <param name="value">The value to look up (not null).</param>
    ''' <param name="valueComparer">The comparer used to compare the values. If it is null, the default value comparer of this enum is used.</param>
    ''' <returns>The enum entry or null if not found.</returns>
    ''' <exception cref="ArgumentNullException">An ArgumentNullException is thrown if <paramref name="value" /> is null.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="The exception is not needed.")> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMemberByValue(ByVal value As TValue, ByVal valueComparer As IEqualityComparer(Of TValue)) As TEnum
        'Handle null
        If (value Is Nothing) Then Throw New ArgumentNullException("value")
        'Get the members
        Dim myMembers As TEnum() = Members
        'Immediately return the member if the values equal
        If (valueComparer Is Nothing) Then
            'Using the default comparer
            For Each myMember As TEnum In myMembers
                Try
                    If (Object.Equals(myMember._Value, value)) Then Return myMember
                Catch
                End Try
            Next
        Else
            'Using the given comparer
            For Each myMember As TEnum In myMembers
                'Immediately return the member if the values equal
                Try
                    If (valueComparer.Equals(myMember._Value, value)) Then Return myMember
                Catch
                End Try
            Next
        End If
        'Return null if not found
        Return Nothing
    End Function

    ''' <summary>Returns a combination with all members that match the given value. If the given value is null, an empty 
    ''' combination is returned. This function uses the value comparer defined by the enumeration (or Object's equal method if 
    ''' not defined) to perform the comparison. This method is thead-safe in a way that if comparing the values throws an 
    ''' exception it is swallowed and ignored. If this is not suitable for you make sure the value and the value-comparer that 
    ''' was possibly provided in the subclasses constructor are thread-safe (or not being manipulated during the function call).</summary>
    ''' <param name="value">The value to look up (not null).</param>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMembersByValues(ByVal value As TValue) As TCombi
        Return Combi.GetInstanceByValuesOptimized(value, ValueComparer)
    End Function

    ''' <summary>Returns a combination with all members that match the given values. Arguments that are null are ignored, null
    ''' values within the array, too. This function uses the value comparer defined by the enumeration (or Object's equal method if 
    ''' not defined) to perform the comparison. This method is thead-safe in a way that if comparing the values throws an 
    ''' exception it is swallowed and ignored. If this is not suitable for you make sure the value and the value-comparer that 
    ''' was possibly provided in the subclasses constructor are thread-safe (or not being manipulated during the function call).</summary>
    ''' <param name="value">The first value to look up.</param>
    ''' <param name="additionalValues">Additional values to look up (split into two parameters for better binding from C#).</param>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMembersByValues(ByVal value As TValue, ByVal ParamArray additionalValues As TValue()) As TCombi
        Return Combi.GetInstanceByValuesOptimized(value, additionalValues, ValueComparer)
    End Function

    ''' <summary>Returns a combination with all members that match the given value. If the given value is null or empty, an 
    ''' empty combination is returned. This function uses the value comparer defined by the enumeration (or Object's equal method 
    ''' if not defined) to perform the comparison. This method is not thread-safe, please ensure that the collection is not
    ''' manipulated during the time it takes to execute this function. If comparing the values throws an exception it is swallowed 
    ''' and ignored. If this is not suitable for you make sure the value and the value-comparer that was possibly provided in the 
    ''' subclasses constructor are thread-safe (or not being manipulated during the function call).</summary>
    ''' <param name="valueCollection">The values to look up.</param>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    Public Shared Function GetMembersByValues(ByVal valueCollection As IEnumerable(Of TValue)) As TCombi
        If (valueCollection Is Nothing) Then Return Combi.Empty
        Dim myArray As TValue() = ToArray(Of TValue)(valueCollection)
        Return Combi.GetInstanceByValuesOptimized(myArray, ValueComparer)
    End Function

    ''' <summary>
    ''' Returns a combination with all members that match the given value. If the given value is null, an empty combination 
    ''' is returned. For comparing the values, the provided comparer is used. If the comparer is null, Object.Equals is used.
    ''' This method is thead-safe in a way that if comparing the values throws an exception it is swallowed and ignored. 
    ''' If this is not suitable for your needs ensure the value and the value-comparer that was possibly provided in the 
    ''' subclasses constructor are thread-safe (or not being manipulated during the function call).
    ''' </summary>
    ''' <param name="value">The value to look up.</param>
    ''' <param name="valueComparer">The comparer used to compare the values. If it is null, Object.Equals is used.</param>
    ''' <returns>The members found as combination</returns>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMembersByValues(ByVal value As TValue, ByVal valueComparer As IEqualityComparer(Of TValue)) As TCombi
        Return Combi.GetInstanceByValuesOptimized(value, valueComparer)
    End Function

    ''' <summary>
    ''' Returns a combination with all members that match the given value. If the given value is null, an empty combination 
    ''' is returned. For comparing the values, the provided comparer is used. If the comparer is null, Object.Equals is used.
    ''' This method is not thread-safe, please ensure that the collection is not manipulated during the time it takes to 
    ''' execute this function. If comparing the values throws an exception it is swallowed and ignored. If this is not suitable 
    ''' for your needs ensure the value and the value-comparer that was possibly provided in the subclasses constructor are 
    ''' thread-safe (or not being manipulated during the function call).
    ''' </summary>
    ''' <param name="valueCollection">The values to look up.</param>
    ''' <param name="valueComparer">The comparer used to compare the values. If it is null, Object.Equals is used.</param>
    ''' <returns>The members found as combination</returns>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from
    ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that
    ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
    ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shared Function GetMembersByValues(ByVal valueCollection As IEnumerable(Of TValue), ByVal valueComparer As IEqualityComparer(Of TValue)) As TCombi
        If (valueCollection Is Nothing) Then Return Combi.Empty
        Dim myArray As TValue() = ToArray(Of TValue)(valueCollection)
        Return Combi.GetInstanceByValuesOptimized(myArray, valueComparer)
    End Function

    'Public Operators

    ''' <summary>Always implicitely allow to cast into the value type (like this is the case with standard 
    ''' .NET enumerations). If the enum value is null, the result is null as well.</summary>
    Public Shared Widening Operator CType(ByVal value As ValueEnum(Of TEnum, TValue, TCombi)) As TValue
        If (value Is Nothing) Then Return Nothing
        Return value.Value
    End Operator

    ''' <summary>
    ''' Values that are defined in this enum may explicitely be casted into the enum. Internally
    ''' this cast operator calls <see cref="GetMemberByValue" /> and if the member was found, it is
    ''' returned, otherwise an <see cref="InvalidCastException" /> thrown.
    ''' </summary>
    ''' <param name="value">The value whos member to look up.</param>
    ''' <returns>The member (or a thrown exception, never null).</returns>
    ''' <exception cref="InvalidCastException">
    ''' The value does not have a corresponding enum member.
    ''' or
    ''' The value is ambiguous and matches more than one enum member.
    ''' </exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if the 
    ''' ValueEnum is incorrectly implemented by the subclass.
    ''' </exception>
    Public Shared Narrowing Operator CType(ByVal value As TValue) As ValueEnum(Of TEnum, TValue, TCombi)
        Dim myResult As Combi = GetMembersByValues(value)
        Dim myMembers As TEnum() = myResult.ToArray()
        Select Case myMembers.Length
            Case 1
                Return myMembers(0)
            Case 0
                Throw New InvalidCastException("The value does not have a corresponding enum member.")
            Case Else
                Throw New InvalidCastException("The value is ambiguous and matches more than one enum member.")
        End Select
    End Operator


    ''' <summary>Combines two enum values to a Combination using a binary OR operation. It is more efficient to use 
    ''' <see cref="GetCombi">GetCombi(..)</see> to combine more than two members. There is no 
    ''' difference between the "OR" and the "+" operators. This operation is thread-safe.</summary>
    ''' <param name="arg1">The first member to combine (null is valid).</param>
    ''' <param name="arg2">The second member to combine (null is valid).</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator Or(ByVal arg1 As ValueEnum(Of TEnum, TValue, TCombi), ByVal arg2 As TEnum) As TCombi
        'Handle null
        If (arg1 Is Nothing) Then
            If (arg2 Is Nothing) Then Return Combi.Empty
            Return arg2.ToCombi()
        End If
        If (arg2 Is Nothing) Then Return arg1.ToCombi()
        'Return new combination
        Dim myArg1 As TEnum = CType(arg1, TEnum)
        Return Combi.GetInstanceOptimized(New TEnum() {myArg1, arg2})
    End Operator

    ''' <summary>Combines two enum values through XOR and returns a new combination instance of the two. If you are combining
    ''' more than two values it is more efficient to initialize an empty combination and the call the Toggle method
    ''' subsequently (or ToggleRange) because each XOR operation allocates a new Combination instance.</summary>
    ''' <param name="arg1">The first member to combine (null is valid).</param>
    ''' <param name="arg2">The second member to combine (null is valid).</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator Xor(ByVal arg1 As ValueEnum(Of TEnum, TValue, TCombi), ByVal arg2 As TEnum) As TCombi
        'Handle null
        If (arg1 Is Nothing) Then
            If (arg2 Is Nothing) Then Return Combi.Empty
            Return arg2.ToCombi()
        End If
        If (arg2 Is Nothing) Then Return arg1.ToCombi()
        'If the members are equal, return an empty combination
        If (arg1 = arg2) Then Return Combi.Empty
        'Otherwise return normal or
        Return (arg1 Or arg2)
    End Operator

    ''' <summary>Binare AND operation of two single members. It returns arg1 if it equals arg2, null otherwise. This operation is thread-safe.</summary>
    ''' <param name="arg1">The first argument, a member (null is valid).</param>
    ''' <param name="arg2">The second argument, a member (null is valid).</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator And(ByVal arg1 As ValueEnum(Of TEnum, TValue, TCombi), ByVal arg2 As TEnum) As TEnum
        'If the members are equal, return a member
        If (arg1 = arg2) Then Return CType(arg1, TEnum)
        'Otherwise return null
        Return Nothing
    End Operator

    ''' <summary>Returns a new instance of Combination that contains all members that are not in arg. This operation is thread-safe.</summary>
    ''' <param name="arg">The argument, a combination of members.</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator Not(ByVal arg As ValueEnum(Of TEnum, TValue, TCombi)) As TCombi
        'Handle null
        If (arg Is Nothing) Then Return Combi.All
        'Return inverted result
        Dim myResult As TCombi = Combi.All
        Return (myResult - CType(arg, TEnum))
    End Operator

    ''' <summary>Binare AND operation of a single member and a combination of members. It returns arg1 if it is contained in arg2, null otherwise. 
    ''' This operation is thread-safe.</summary>
    ''' <param name="arg1">The first argument, a member.</param>
    ''' <param name="arg2">The second argument, a combination of members.</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator And(ByVal arg1 As ValueEnum(Of TEnum, TValue, TCombi), ByVal arg2 As TCombi) As TEnum
        'Handle null
        If (arg1 Is Nothing) OrElse (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return Nothing
        'Merge both flags
        Dim myArg1 As TEnum = DirectCast(arg1, TEnum)
        If (arg2.Contains(myArg1)) Then Return myArg1
        Return Nothing
    End Operator

    ''' <summary>Binare AND operation of a single member and a combination of members. It returns arg2 if it is contained in arg1, null otherwise. 
    ''' This operation is thread-safe.</summary>
    ''' <param name="arg1">The first argument, a combination of members.</param>
    ''' <param name="arg2">The second argument, a member.</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator And(ByVal arg1 As TCombi, ByVal arg2 As ValueEnum(Of TEnum, TValue, TCombi)) As TEnum
        Return (arg2 And arg1)
    End Operator

    ''' <summary>Subtracts a member from a member (null is returned if arg2 equlas arg1, otherwise arg1 is returned).
    ''' This operation is thread-safe.</summary>
    ''' <param name="arg1">The first argument, a member.</param>
    ''' <param name="arg2">The argument to subtract, a member.</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator -(ByVal arg1 As ValueEnum(Of TEnum, TValue, TCombi), ByVal arg2 As TEnum) As TEnum
        If (arg1 Is Nothing) Then Return Nothing
        If (arg2 Is Nothing) Then Return CType(arg1, TEnum)
        If (arg1.Equals(arg2)) Then Return Nothing
        Return CType(arg1, TEnum)
    End Operator

    ''' <summary>Subtracts multiple members from a member (null is returned if arg2 contains arg1, otherwise arg1 is returned).
    ''' This operation is thread-safe.</summary>
    ''' <param name="arg1">The first argument, a member.</param>
    ''' <param name="arg2">The argument to subtract, a combination of members.</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator -(ByVal arg1 As ValueEnum(Of TEnum, TValue, TCombi), ByVal arg2 As TCombi) As TEnum
        'Handle null
        If (arg1 Is Nothing) Then Return Nothing
        Dim myArg1 As TEnum = CType(arg1, TEnum)
        If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return myArg1
        'Handle combination
        If (arg2.Contains(myArg1)) Then
            Return Nothing
        End If
        Return myArg1
    End Operator

    ''' <summary>Combines two enum values to a Combination using a binary OR operation. It is more efficient to use 
    ''' <see cref="GetCombi">GetCombi(..)</see> to combine more than two members. There is no 
    ''' difference between the "+" and the "OR" operators. This operation is thread-safe.</summary>
    ''' <param name="arg1">The first member to combine (null is valid).</param>
    ''' <param name="arg2">The second member to combine (null is valid).</param>
    <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
    Public Shared Operator +(ByVal arg1 As ValueEnum(Of TEnum, TValue, TCombi), ByVal arg2 As TEnum) As TCombi
        Return (arg1 Or arg2)
    End Operator

    ''' <summary>Compares two member and returns true if they are equal. This operation is thead-safe.</summary>
    ''' <param name="arg1">The first argument, a member (null is valid).</param>
    ''' <param name="arg2">The second argument, a member (null is valid).</param>
    Public Shared Operator =(ByVal arg1 As ValueEnum(Of TEnum, TValue, TCombi), ByVal arg2 As TEnum) As Boolean
        If (Object.ReferenceEquals(arg1, arg2)) Then Return True 'if both are null, true is returned
        If (arg1 Is Nothing) OrElse (arg2 Is Nothing) Then Return False
        Return arg1.Equals(arg2)
    End Operator

    ''' <summary>Compares a member with a value and returns true if they are equal. This operation is thead-safe if the value instance
    ''' and the value comparer are thread-safe (or not manipulated by other threads).</summary>
    ''' <param name="arg1">The first argument, a member.</param>
    ''' <param name="arg2">The second argument, a value.</param>
    Public Shared Operator =(ByVal arg1 As ValueEnum(Of TEnum, TValue, TCombi), ByVal arg2 As TValue) As Boolean
        'Check null
        If (arg1 Is Nothing) Then
            If (arg2 Is Nothing) Then Return True
            Return False
        End If
        'Return whether x equals y
        Return (arg1.Equals(arg2))
    End Operator

    ''' <summary>Compares a member with a value and returns true if they are equal. This operation is thead-safe if the value instance
    ''' and the value comparer are thread-safe (or not manipulated by other threads).</summary>
    ''' <param name="arg1">The first argument, a value.</param>
    ''' <param name="arg2">The second argument, a member.</param>
    Public Shared Operator =(ByVal arg1 As TValue, ByVal arg2 As ValueEnum(Of TEnum, TValue, TCombi)) As Boolean
        Return (arg2 = arg1)
    End Operator


    ''' <summary>Compares two member and returns true if they are inequal.</summary>
    ''' <param name="arg1">The first argument, a member (null is valid).</param>
    ''' <param name="arg2">The second argument, a member (null is valid).</param>
    Public Shared Operator <>(ByVal arg1 As ValueEnum(Of TEnum, TValue, TCombi), ByVal arg2 As TEnum) As Boolean
        Return Not (arg1 = arg2)
    End Operator

    ''' <summary>Compares a member with a value and returns true if they are inequal. This operation is thead-safe if the value instance
    ''' and the value comparer are thread-safe (or not manipulated by other threads).</summary>
    ''' <param name="arg1">The first argument, a member.</param>
    ''' <param name="arg2">The second argument, a value.</param>
    Public Shared Operator <>(ByVal arg1 As ValueEnum(Of TEnum, TValue, TCombi), ByVal arg2 As TValue) As Boolean
        Return (Not (arg1 = arg2))
    End Operator

    ''' <summary>Compares a member with a value and returns true if they are inequal. This operation is thead-safe if the value instance
    ''' and the value comparer are thread-safe (or not manipulated by other threads).</summary>
    ''' <param name="arg1">The first argument, a value.</param>
    ''' <param name="arg2">The second argument, a member.</param>
    Public Shared Operator <>(ByVal arg1 As TValue, ByVal arg2 As ValueEnum(Of TEnum, TValue, TCombi)) As Boolean
        Return (arg2 <> arg1)
    End Operator

    'Framework Properties

    ''' <summary>Returns true if the implementation of this CustomEnum looks correct.</summary>
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="The exception is not needed.")> _
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared ReadOnly Property IsValid() As Boolean
        Get
            'If the initialization does not throw an exception, return true
            Try
                InitAndReturnMembers()
                Return True
            Catch
            End Try
            'Otherwise return false
            Return False
        End Get
    End Property

    'Framework Methods

    ''' <summary>Returns a combination with this member set. This method is optimized and returns always the same instance.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    Public Function ToCombi() As TCombi
        Dim myResult As TCombi = _Combination
        If (myResult Is Nothing) Then
            myResult = Combi.GetInstanceRaw(CType(Me, TEnum))
            _Combination = myResult
        End If
        Return myResult
    End Function

    ''' <summary>Returns the name of the enum.</summary>
    Public Overrides Function ToString() As String
        Dim myResult As String = _Name
        If (myResult Is Nothing) Then
            Try
                Return Name 'initializes the name if possible
            Catch ex As InvalidOperationException
                Return "[not initialized]"
            End Try
        End If
        Return myResult
    End Function

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="The exception is not needed.")> _
    Private ReadOnly Property DebuggerDisplayValue As String
        Get
            Dim myResult As New StringBuilder()
            'Get the name
            Dim myName As String = _Name
            If (myName Is Nothing) Then
                Try
                    myName = Name
                Catch ex As InvalidOperationException
                End Try
            End If
            'Add the name
            If (myName Is Nothing) Then
                myResult.Append("[unknown]")
            Else
                myResult.Append(myName)
            End If
            'Add the value (cropped to 50 chars)
            Dim myValue As String = Nothing
            Try
                myValue = Value.ToString()
                If (myValue IsNot Nothing) Then
                    If (myValue.Length > 50) Then myValue = myValue.Substring(0, 47) & "..."
                    myResult.Append(" (")
                    myResult.Append(myValue)
                    myResult.Append(")"c)
                End If
            Catch
            End Try
            'Return the result
            Return myResult.ToString()
        End Get
    End Property

    ''' <summary>Returns true if the given object is a member and equals this member (see according Equals overload), 
    ''' or if it is of type TValue and equals this instance using the member's ValueComparer (see according Equals overload), 
    ''' or if it is of type Combination and contains only this instance as member (see according Equals overload).</summary>
    ''' <param name="obj">The other instance.</param>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        'Handle null
        If (obj Is Nothing) Then Return False
        'Handle Member
        Dim myMember As TEnum = TryCast(obj, TEnum)
        If (myMember IsNot Nothing) Then
            Return Equals(myMember)
        End If
        'Handle Value
        If (TypeOf obj Is TValue) Then
            Dim myValue As TValue = DirectCast(obj, TValue)
            Return Equals(myValue)
        End If
        'Handle Combination
        Dim myCombination As Combi = TryCast(obj, Combi)
        If (myCombination IsNot Nothing) Then
            Return Equals(myCombination)
        End If
        'Otherwise return false
        Return False
    End Function

    ''' <summary>By default, returns true if the other instance is the same reference as this one, false otherwise. 
    ''' This behavior may be overwritten in the subclass, eg. if there are two defined values that are equal.</summary>
    ''' <param name="other">The other instance.</param>
    Public Overridable Overloads Function Equals(ByVal other As TEnum) As Boolean
        If (Object.ReferenceEquals(Me, other)) Then Return True
        Return False
    End Function

    ''' <summary>Returns true if the combination contains this instance as the only member.</summary>
    ''' <param name="other">The other instance.</param>
    Public Overloads Function Equals(ByVal other As Combi) As Boolean
        If (other Is Nothing) Then Return False
        Return other.Equals(Me)
    End Function

    ''' <summary>Compares this member with a value and returns true if they match. There is also an equal operator 
    ''' that is doing the same if you preferre to work with operators.</summary>
    ''' <param name="other">The value to compare with this member.</param>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Overridable Overloads Function Equals(ByVal other As TValue) As Boolean Implements System.IEquatable(Of TValue).Equals
        If (other Is Nothing) Then Return False
        Dim myComparer As IEqualityComparer(Of TValue) = ValueComparer
        If (myComparer Is Nothing) Then
            Return Object.Equals(Me._Value, other)
        End If
        Return myComparer.Equals(Me._Value, other)
    End Function

    ''' <summary>Returns the hashcode of the value to ensure member and value have the same hashcodes.</summary>
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Overrides Function GetHashCode() As Int32
        Return _Value.GetHashCode()
    End Function

    'Private Properties

    ''' <summary>Returns all members of this enum. The first time this property is called they are evaluated through reflection 
    ''' and then cached in a static variable. Watch out: Do not call this function from within the subclasses constructor!</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if a property or method is accessed from 
    ''' within the subclasses constructor that calls this property und would cause the member array to be initialized (which is not 
    ''' possible to that time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function 
    ''' is called later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
    ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
    ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
    ''' to write name and index back).</exception>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared ReadOnly Property Members As TEnum()
        Get
            Dim myResult As TEnum() = _Members
            If (myResult Is Nothing) Then
                myResult = InitAndReturnMembers()
            End If
            Return myResult
        End Get
    End Property

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared ReadOnly Property NameComparer As IEqualityComparer(Of String)
        Get
            Dim myResult As IEqualityComparer(Of String) = _NameComparer
            If (myResult Is Nothing) Then
                myResult = InitAndReturnNameComparer(Members)
            End If
            Return myResult
        End Get
    End Property

    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this property causes the member 
    ''' array to be initialized but there are some design errors in the subclass: The subclass may not assign the same instance 
    ''' to multiple fields (because Name and Index are written back to each instance the information from the previous field 
    ''' would be overwritten) or if the instance is null (which also makes it difficult to write name and index back).</exception>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
    Private Shared ReadOnly Property ValueComparer As IEqualityComparer(Of TValue)
        Get
            'Initialize the value comparer if not yet initialized (that was an ugly bug)
            If (_IsFirstInstance) Then
                InitAndReturnMembers()
            End If
            'Return the value comparer (that usually is null)
            Return _ValueComparer
        End Get
    End Property

    ''' <summary>For debugging only. Hides the static infos from the instance but still allows to browse it if needed.</summary>
    Private Shared ReadOnly Property Zzzzz As StaticFields
        Get
            Return StaticFields.Singleton
        End Get
    End Property

    'Private Methods

    ''' <summary>Called after the members of this CustomEnum have been initialized. This method does nothing by
    ''' default but may be overwritten in the subclass.</summary>
    Protected Overridable Sub OnMemberInitialized()
    End Sub

    'Private Functions

    Private Shared Function GetMembersByNames(ByVal dict As Dictionary(Of String, String)) As TCombi
        Dim myMembers As TEnum() = Members
        Dim myResult As New List(Of TEnum)(dict.Count)
        'Fill in the members
        For Each myMember As TEnum In myMembers
            If (dict.Remove(myMember._Name)) Then
                myResult.Add(myMember)
                If (dict.Count = 0) Then Exit For
            End If
        Next
        'Return the members as combination
        Select Case myResult.Count
            Case 0
                Return Combi.Empty
            Case 1
                Return myResult(0).ToCombi()
        End Select
        'Return the result
        Return Combi.GetInstanceOptimized(myResult.ToArray())
    End Function

    ''' <summary>.NET 2.0 support.</summary>
    Private Shared Function ToArray(Of T)(ByVal collection As IEnumerable(Of T)) As T()
        'Handle null
        If (collection Is Nothing) Then Return New T() {}
        'Handle array
        Dim myArray As T() = TryCast(collection, T())
        If (myArray IsNot Nothing) Then Return myArray
        'Initialize result list
        Dim myResult As List(Of T)
        Dim myCollection As ICollection(Of T) = TryCast(collection, ICollection(Of T))
        If (myCollection Is Nothing) Then
            myResult = New List(Of T)()
        Else
            myResult = New List(Of T)(myCollection.Count)
        End If
        'Fill in elements
        For Each myElement As T In collection
            myResult.Add(myElement)
        Next
        'Return result
        Return myResult.ToArray()
    End Function

    ''' <summary>Initializes the members if they are not yet initialized.</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if any property or method
    ''' is accessed from within the subclasses constructor that would cause the member array to be initialized (that
    ''' is not possible to that time because the fields are not yet assigned). An InvalidOperationException is also
    ''' thrown if the same instance is assigned to multiple static fields (that is not possible because Name and Index
    ''' are set to the instance)</exception>
    ''' <returns>A reference to the initialized array.</returns>
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="CustomEnum")> _
    Private Shared Function InitAndReturnMembers() As TEnum()
        Dim myResult As TEnum() = _Members
        If (myResult IsNot Nothing) Then Return myResult
        'Initialize the members
        SyncLock (_Lock)
            myResult = _Members
            If (myResult IsNot Nothing) Then Return myResult
            myResult = PrivateGetMembers()
            If (myResult Is Nothing) Then Throw New InvalidOperationException("Internal error in CustomEnum.")
            'Init comparer and perfom a duplicate check
            If (_NameComparer Is Nothing) Then
                InitAndReturnNameComparer(myResult)
            End If
            'Tell the instance it is initialized
            For Each myMember As TEnum In myResult
                myMember.OnMemberInitialized()
            Next
            'Assign the members (do not assign before the initialization has completed, let other theads wait)
            _Members = myResult
        End SyncLock
        'Return the result
        Return myResult
    End Function

    ''' <summary>Initializes and returns the string comparer used to compare the names.</summary>
    ''' <param name="memberArray">A reference to the member array. Because during initialization it is not yet assigned to _Members it is passed along.</param>
    ''' <exception cref="InvalidOperationException">An invalid operation exception is thrown if there are ambiguous names (not easy to produce in VB).</exception>
    Private Shared Function InitAndReturnNameComparer(ByVal memberArray As TEnum()) As IEqualityComparer(Of String)
        Dim myComparer As IEqualityComparer(Of String) = _NameComparer
        If (myComparer IsNot Nothing) Then Return myComparer
        SyncLock (_Lock)
            myComparer = _NameComparer
            If (myComparer IsNot Nothing) Then Return myComparer
            'Determine the comparer
            myComparer = If(InitAndReturnCaseSensitive(memberArray), StringComparer.Ordinal, StringComparer.OrdinalIgnoreCase)
            'Check for duplicates (happens if the constructor explicitely sets the case-insensitive flag but has two fields that differ only by case,
            'or if the enum has multiple hierarchical subclasses and the overwritten properties do not hide the parent ones (like this is the case 
            'in JScript.NET).
            If (HasDuplicateNames(memberArray, myComparer)) Then
                Throw New InvalidOperationException("Internal error in " & GetType(TEnum).Name & ", the member names are ambiguous.")
            End If
            'If everything is okay, assign the comparer
            _NameComparer = myComparer
            'And return it
            Return myComparer
        End SyncLock
    End Function

    ''' <summary>Gets the members (during initialization).</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if any property or method
    ''' is accessed from within the subclasses constructor that would cause the member array to be initialized (that
    ''' is not possible to that time because the fields are not yet assigned). An InvalidOperationException is also
    ''' thrown if the same instance is assigned to multiple static fields (that is not possible because Name and Index
    ''' are set to the instance)</exception>
    Private Shared Function PrivateGetMembers() As TEnum()
        Dim myList As New List(Of TEnum)
        AddFields(myList)
        AddGetters(myList)
        Return myList.ToArray()
    End Function

    ''' <summary>Adds all public static readonly fields that are of type TEnum (flat, only of class TEnum).</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if any property or method
    ''' is accessed from within the subclasses constructor that would cause the member array to be initialized (that
    ''' is not possible to that time because the fields are not yet assigned). An InvalidOperationException is also
    ''' thrown if the same instance is assigned to multiple static fields (that is not possible because Name and Index
    ''' are set to the instance)</exception>
    Private Shared Sub AddFields(ByVal aList As List(Of TEnum))
        Dim myFlags As BindingFlags = (BindingFlags.Static Or BindingFlags.Public)
        Dim myFields As FieldInfo() = GetType(TEnum).GetFields(myFlags)
        For Each myField As FieldInfo In myFields
            'Ignore fields of other types
            If (Not (GetType(TEnum).IsAssignableFrom(myField.FieldType))) Then Continue For
            'Ignore read/write fields
            If (Not myField.IsInitOnly) Then Continue For
            'Ignore flagged fields
            If (IsFlaggedToIgnore(myField)) Then Continue For
            'Add field
            Dim myEntry As TEnum = CType(myField.GetValue(Nothing), TEnum)
            AddEntry(myEntry, myField.Name, aList)
        Next
    End Sub

    ''' <summary>Adds all public static getter properties without indexer that are of type TEnum (flat, only of class TEnum).</summary>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if any property or method
    ''' is accessed from within the subclasses constructor that would cause the member array to be initialized (that
    ''' is not possible to that time because the fields are not yet assigned). An InvalidOperationException is also
    ''' thrown if the same instance is assigned to multiple static fields (that is not possible because Name and Index
    ''' are set to the instance)</exception>
    Private Shared Sub AddGetters(ByVal aList As List(Of TEnum))
        Dim myFlags As BindingFlags = (BindingFlags.Static Or BindingFlags.Public Or BindingFlags.GetProperty)
        Dim myProperties As PropertyInfo() = GetType(TEnum).GetProperties(myFlags)
        For Each myProperty As PropertyInfo In myProperties
            'Ignore properties of other types
            If (Not (GetType(TEnum).IsAssignableFrom(myProperty.PropertyType))) Then Continue For
            'Look only at read-only properties
            If (myProperty.CanWrite) Then Continue For
            If (Not myProperty.CanRead) Then Continue For
            'Ignore indexed properties
            If (myProperty.GetIndexParameters().Length > 0) Then Continue For
            'Ignore flagged properties
            If (IsFlaggedToIgnore(myProperty)) Then Continue For
            'Invoke the property twice and check whether the same instance is returned (it is a requirement)
            Dim myEntry As TEnum = CType(myProperty.GetValue(Nothing, Nothing), TEnum)
            Dim myEntry2 As TEnum = CType(myProperty.GetValue(Nothing, Nothing), TEnum)
            If (Not Object.ReferenceEquals(myEntry, myEntry2)) Then
                Throw New InvalidOperationException("Internal error in " & GetType(TEnum).Name & "! Property " & myProperty.Name & " returned different instances when invoked multiple times. Ensure always the same instance is returned.")
            End If
            'Add the entry
            AddEntry(myEntry, myProperty.Name, aList)
        Next
    End Sub

    Private Shared Function IsFlaggedToIgnore(ByVal field As FieldInfo) As Boolean
        Return IsFlaggedToIgnore(field.GetCustomAttributes(False))
    End Function

    Private Shared Function IsFlaggedToIgnore(ByVal [property] As PropertyInfo) As Boolean
        Return IsFlaggedToIgnore([property].GetCustomAttributes(False))
    End Function

    Private Shared Function IsFlaggedToIgnore(ByVal attributes As Object()) As Boolean
        If (attributes Is Nothing) OrElse (attributes.Length = 0) Then Return False
        For Each myInstance As Object In attributes
            If (myInstance Is Nothing) Then Continue For
            If (TypeOf myInstance Is CustomEnumIgnoreAttribute) Then Return True
        Next
        Return False
    End Function


    ''' <summary>Adds an entry to the list. During this process it is checked whether the same instance was already added (throws
    ''' an InvalidOperationException) and also assigns the member's name and index.</summary>
    ''' <param name="member">The member to add.</param>
    ''' <param name="name">The field or property name where the member was assigned to.</param>
    ''' <param name="result">A list of members to which the values are added.</param>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if any property or method
    ''' is accessed from within the subclasses constructor that would cause the member array to be initialized (that
    ''' is not possible to that time because the fields are not yet assigned). An InvalidOperationException is also
    ''' thrown if the same instance is assigned to multiple static fields (that is not possible because Name and Index
    ''' are set to the instance)</exception>
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="CustomEnumIgnoreAttribute", Justification:="That's the name of the attribute!")> _
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="readonly")> _
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="OnMemberInitialized")> _
    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="GetMemberBy")> _
    Private Shared Sub AddEntry(ByVal member As TEnum, ByVal name As String, ByVal result As List(Of TEnum))
        'Check for instance conflicts
        If (member Is Nothing) Then
            Throw New InvalidOperationException("Internal error in " & GetType(TEnum).Name & "! Do not access any properties (like property Name and Index) nor methods (e.g. all ""GetMemberBy..."") that trigger the initialization of the member array from within the subclasses constructor as the fields are not yet initialized at that time! Also avoid to provide any public static readonly fields of type " & GetType(TEnum).Name & " that have a null-value. You can override method OnMemberInitialized if you need to do some post-initialization, or change the static fields into static getter properties with lazy initialization.")
        End If
        If (member._Name IsNot Nothing) Then
            Throw New InvalidOperationException("Internal error in " & GetType(TEnum).Name & "! It's invalid to assign the same instance to multiple fields/properties that are treated as members (a conflict arises when assigning name and index to the instance). You can use the CustomEnumIgnoreAttribute to flag a field/property to not be treated as a member declaration but as an additional custom field/property of the class.")
        End If
        'Set the name and index
        member._Name = name
        member._Index = result.Count
        'Add to the list
        result.Add(member)
    End Sub

    ''' <summary>Determines whether case-sensitive name comparison is needed (two or more members differ only by name, 
    ''' e.g. when the subclass was written in C#) or not.</summary>
    ''' <param name="memberArray">A reference to the member array. Because during initialization it is not yet assigned to _Members it is passed along.</param>
    ''' <returns>True if case sensitive comparison is needed.</returns>
    Private Shared Function InitAndReturnCaseSensitive(ByVal memberArray As TEnum()) As Boolean
        Dim myResult As Boolean? = _CaseSensitive
        If (myResult Is Nothing) Then
            myResult = HasDuplicateNames(memberArray, StringComparer.OrdinalIgnoreCase)
            _CaseSensitive = myResult
        End If
        Return myResult.Value
    End Function

    ''' <summary>Determines whether there are duplicate names when compared with the given comparer.</summary>
    ''' <param name="memberArray">A reference to the member array. Because during initialization it is not yet assigned to _Members it is passed along.</param>
    ''' <param name="comparer">An equality comparer used to comparer the names.</param>
    ''' <returns>True if there were duplicate names.</returns>
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="The exception is not needed.")> _
    Private Shared Function HasDuplicateNames(ByVal memberArray As TEnum(), ByVal comparer As IEqualityComparer(Of String)) As Boolean
        Dim myDict As New Dictionary(Of String, TEnum)(comparer)
        Try
            For Each myEntry As TEnum In memberArray
                myDict.Add(myEntry._Name, myEntry)
            Next
        Catch
            Return True
        End Try
        Return False
    End Function

    <SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId:="MustInherit")> _
    Private Shared Sub CheckConstructors()
        'Check whether we are granted permission to access private members through reflection
        Dim permission As New ReflectionPermission(PermissionState.Unrestricted)
        Try
            permission.Demand()
        Catch ex As SecurityException
            Return
        End Try
        'Get all constructors
        Dim myConstructors As ConstructorInfo() = GetType(TEnum).GetConstructors(BindingFlags.CreateInstance Or BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
        For Each myConstructor As ConstructorInfo In myConstructors
            'Should be private
            If (myConstructor.IsPrivate) Then Continue For
            If (myConstructor.DeclaringType.IsAbstract) Then
                If (myConstructor.IsFamily) Then Continue For
                If (myConstructor.IsFamilyAndAssembly) Then Continue For
            End If
            'Notify implementation error
            Throw New InvalidOperationException("All constructors of " & GetType(TEnum).Name & " must be declared private. If the class is defined as ""MustInherit"", protected constructors are tolerated!")
        Next
    End Sub


    '**********************************************************************
    ' INNER CLASS: StaticInfo
    '**********************************************************************

    ''' <summary>For debugger only. Allowes to hide static information from normal instances but the information is still 
    ''' browsable if needed.</summary>
    <DebuggerDisplay("Expand for static infos...", Name:="(static)")> _
    Private Class StaticFields

        'Private Fields
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private Shared _Singleton As StaticFields

        'Constructors

        Private Sub New()
        End Sub

        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Public Shared ReadOnly Property Singleton As StaticFields
            Get
                Dim myResult As StaticFields = _Singleton
                If (myResult Is Nothing) Then
                    myResult = New StaticFields()
                    _Singleton = myResult
                End If
                Return myResult
            End Get
        End Property

        'Public Properties

        Public Shared ReadOnly Property Members As TEnum()
            Get
                Return ValueEnum(Of TEnum, TValue, TCombi).Members
            End Get
        End Property

        Public Shared ReadOnly Property First As TEnum
            Get
                Return ValueEnum(Of TEnum, TValue, TCombi).First
            End Get
        End Property

        Public Shared ReadOnly Property Last As TEnum
            Get
                Return ValueEnum(Of TEnum, TValue, TCombi).Last
            End Get
        End Property

        Public Shared ReadOnly Property CaseSensitive As Boolean
            Get
                Return ValueEnum(Of TEnum, TValue, TCombi).CaseSensitive
            End Get
        End Property

        Public Shared ReadOnly Property IsValid As Boolean
            Get
                Return ValueEnum(Of TEnum, TValue, TCombi).IsValid
            End Get
        End Property

        Public Shared ReadOnly Property NameComparer As IEqualityComparer(Of String)
            Get
                Return ValueEnum(Of TEnum, TValue, TCombi).NameComparer
            End Get
        End Property

        Public Shared ReadOnly Property ValueComparer As IEqualityComparer(Of TValue)
            Get
                Return ValueEnum(Of TEnum, TValue, TCombi).ValueComparer
            End Get
        End Property

        Public Shared ReadOnly Property ValueType As Type
            Get
                Return ValueEnum(Of TEnum, TValue, TCombi).ValueType
            End Get
        End Property

        Public Shared ReadOnly Property CombiType As Type
            Get
                Return ValueEnum(Of TEnum, TValue, TCombi).CombiType
            End Get
        End Property

    End Class


    '**************************************************************************
    ' INNER CLASS: Combi
    '**************************************************************************

    ''' <summary>Class that provides support for combining multiple enumeration values, similair to standard enums in combination
    ''' with the <see cref="FlagsAttribute">Flags</see> attribute. A combination is mutually readonly but there are many 
    ''' operators and method that allow to create new combinations. This class aims to be thead-safe. Because of that, it does not 
    ''' support IEnumerable(Of TEnum) directly as arguments for methods and operators. But there is a constructor that takes an 
    ''' IEnumerable(Of TEnum) and like this you can convert your collection into a Combination and then do whatever you like, 
    ''' thead-safe.</summary>
    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix", Justification:="It is no collection and I like it like that.")> _
    <SuppressMessage("Microsoft.Design", "CA1034:NestedTypesShouldNotBeVisible", Justification:="We need very much, that is a feature, not a bug!")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    <DebuggerDisplay("{DebuggerDisplayValue}")> _
    Public Class Combi
        Implements IEnumerable(Of TEnum)
        Implements IEquatable(Of Combi)
        Implements IEquatable(Of TEnum)

        'Private Fields
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private Shared _Empty As TCombi = Nothing
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private Shared _All As TCombi = Nothing
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private _Flags As Boolean()
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private _Count As Int32 = -1
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private _HashCode As Int32? = Nothing

        'Constructors

        ''' <summary>Creates a new empty Combination instance. Do not call this constructor, use Combi.Empty if you need 
        ''' a reference to an empty combination, or use an operator overload or a GetCombi-method to get a combination containing
        ''' members.</summary>
        Public Sub New()
        End Sub

        ''' <summary>Returns an instance of combination containing the given member. This function is thread-safe.</summary>
        <EditorBrowsable(EditorBrowsableState.Never)> _
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Public Shared Function GetInstanceOptimized(ByVal member As TEnum) As TCombi
            If (member Is Nothing) Then Return Empty
            Return member.ToCombi()
        End Function

        ''' <summary>Returns an instance of combination containing the given members. This function is thread-safe.</summary>
        <EditorBrowsable(EditorBrowsableState.Never)> _
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Public Shared Function GetInstanceOptimized(ByVal member1 As TEnum, ByVal member2 As TEnum()) As TCombi
            'Handle empty
            If ((member2 Is Nothing) OrElse (member2.Length = 0)) Then
                If (member1 Is Nothing) Then Return Empty
                Return member1.ToCombi()
            End If
            'Determine result
            Dim myResult As New TCombi()
            myResult.Set(member1)
            myResult.Set(member2)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Function

        ''' <summary>Returns <see cref="Combi.Empty" /> if null or empty, otherwise an instance containing the given members is returned. 
        ''' This function is thread-safe.</summary>
        ''' <param name="members">The members to get.</param>
        <EditorBrowsable(EditorBrowsableState.Never)> _
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Public Shared Function GetInstanceOptimized(ByVal members As TEnum()) As TCombi
            'Determine result
            Dim myResult As New TCombi()
            myResult.Set(members)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Function

        ''' <summary>Returns <see cref="Combi.Empty" /> if null or empty, otherwise the same instance as provided as argument is returned. This 
        ''' function is thread-safe.</summary>
        ''' <param name="members">The members to get.</param>
        <EditorBrowsable(EditorBrowsableState.Never)> _
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Public Shared Function GetInstanceOptimized(ByVal members As Combi) As TCombi
            If (members Is Nothing) OrElse (members.IsEmpty) Then Return Empty
            Return CType(members, TCombi)
        End Function

        'Watch out: Always return a new instance from GetInstanceRaw-functions, never return the one that was given as a parameter,
        '           don't return Combi.Empty nor Combi.All as the instances may still be manipulated after the call.
        '           This is the only place where this rule applies, in all other places it's save and recommended to recycle 
        '           instances as they are immutable as soon as they leave this class.

        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Private Shared Function GetInstanceRaw() As TCombi
            Return New TCombi()
        End Function

        ''' <summary>Constructs a new combination and initializes the given member.</summary>
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <EditorBrowsable(EditorBrowsableState.Never)> _
        Public Shared Function GetInstanceRaw(ByVal member As TEnum) As TCombi
            Dim myResult As New TCombi()
            myResult.Set(member)
            Return myResult
        End Function

        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Private Shared Function GetInstanceRaw(ByVal members As Combi) As TCombi
            Dim myResult As New TCombi()
            If (members Is Nothing) OrElse (members.Count = 0) Then Return myResult
            Array.Copy(members.Flags, myResult.Flags, members.Flags.Length)
            Return myResult
        End Function


        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <EditorBrowsable(EditorBrowsableState.Never)> _
        Public Shared Function GetInstanceByIndexesOptimized(ByVal indexArray As Int32()) As TCombi
            'Handle empty
            If (indexArray Is Nothing) OrElse (indexArray.Length = 0) Then Return Empty
            'Determine result
            Dim myResult As New TCombi()
            Dim myFlags As Boolean() = myResult.Flags
            Dim myLength As Int32 = myFlags.Length
            For Each myIndex As Int32 In indexArray
                If (myIndex >= myLength) OrElse (myIndex < 0) Then Continue For
                myFlags(myIndex) = True
            Next
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return the result
            Return myResult
        End Function

        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <EditorBrowsable(EditorBrowsableState.Never)> _
        Public Shared Function GetInstanceByIndexesOptimized(ByVal index As Int32, ByVal indexArray As Int32()) As TCombi
            'Determine result
            Dim myResult As New TCombi()
            Dim myFlags As Boolean() = myResult.Flags
            Dim myLength As Int32 = myFlags.Length
            'Add single index
            If (index < myLength) AndAlso (index > -1) Then
                myFlags(index) = True
            End If
            'Add array
            If (indexArray IsNot Nothing) AndAlso (indexArray.Length > 0) Then
                For Each myIndex As Int32 In indexArray
                    If (myIndex >= myLength) OrElse (myIndex < 0) Then Continue For
                    myFlags(myIndex) = True
                Next
            End If
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return the result
            Return myResult
        End Function

        ''' <summary>Is exposed through TEnum.GetMembersByValue, but because of private field access must be implemented here.</summary>
        ''' <param name="value">The value to look up (null is ignored).</param>
        ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this function is accessed from 
        ''' within the subclasses constructor what would cause the member array to be initialized (which is not possible to that 
        ''' time because the fields are not yet assigned). An InvalidOperationException is also thrown if this function is called
        ''' later and it causes the member array to be initialized but there are some design errors in the subclass: The subclass 
        ''' may not assign the same instance to multiple fields (because Name and Index are written back to each instance the
        ''' information from the previous field would be overwritten) or if the instance is null (which also makes it difficult
        ''' to write name and index back).</exception>
        <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="To avoid multi-threading issues, any exception may be thrown.")> _
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <EditorBrowsable(EditorBrowsableState.Never)> _
        Public Shared Function GetInstanceByValuesOptimized(ByVal value As TValue, ByVal valueComparer As IEqualityComparer(Of TValue)) As TCombi
            'Handle null
            If (value Is Nothing) Then Return Empty
            'Get the members
            Dim myMembers As TEnum() = Members
            Dim myResult As New TCombi()
            Dim myFlags As Boolean() = myResult.Flags
            'Use Object.Equals to compare
            If (valueComparer Is Nothing) Then
                For i As Int32 = 0 To myMembers.Length - 1
                    Dim myMember As TEnum = myMembers(i)
                    Try
                        If (Object.Equals(myMember._Value, value)) Then
                            myFlags(i) = True
                        End If
                    Catch
                    End Try
                Next
            Else
                'Using the given comparer
                For i As Int32 = 0 To myMembers.Length - 1
                    Dim myMember As TEnum = myMembers(i)
                    Try
                        If (valueComparer.Equals(myMember._Value, value)) Then
                            myFlags(i) = True
                        End If
                    Catch
                    End Try
                Next
            End If
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return the combination
            Return myResult
        End Function

        ''' <summary>Gets a combination instance that contains the members specified by the given values.</summary>
        ''' <param name="values">The value that define the selected members.</param>
        ''' <param name="valueComparer">The comparer used for the value comparison.</param>
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <EditorBrowsable(EditorBrowsableState.Never)> _
        Public Shared Function GetInstanceByValuesOptimized(ByVal values As TValue(), ByVal valueComparer As IEqualityComparer(Of TValue)) As TCombi
            'Handle null
            If (values Is Nothing) OrElse (values.Length = 0) Then Return Empty
            If (values.Length = 1) Then Return GetInstanceByValuesOptimized(values(0), valueComparer)
            'Convert values and comparer into a dictionary
            Dim myDict As Dictionary(Of TValue, TValue)
            If (valueComparer Is Nothing) Then
                myDict = New Dictionary(Of TValue, TValue)()
            Else
                myDict = New Dictionary(Of TValue, TValue)(valueComparer)
            End If
            For Each myValue As TValue In values
                If (myValue Is Nothing) Then Continue For
                myDict.Item(myValue) = myValue
            Next
            'Return the combination
            Return GetInstanceByValuesOptimized(myDict)
        End Function

        ''' <summary>Gets a combination instance that contains the members specified by the given values.</summary>
        ''' <param name="value">The value that define the first member to select.</param>
        ''' <param name="additionalValues">The additional values that define the members to select.</param>
        ''' <param name="valueComparer">The comparer used for the value comparison.</param>
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <EditorBrowsable(EditorBrowsableState.Never)> _
        Public Shared Function GetInstanceByValuesOptimized(ByVal value As TValue, ByVal additionalValues As TValue(), ByVal valueComparer As IEqualityComparer(Of TValue)) As TCombi
            'Handle null
            If (additionalValues Is Nothing) OrElse (additionalValues.Length = 0) Then
                If (value Is Nothing) Then Return Empty
                Return GetInstanceByValuesOptimized(value, valueComparer)
            End If
            'Convert values and comparer into a dictionary
            Dim myDict As Dictionary(Of TValue, TValue)
            If (valueComparer Is Nothing) Then
                myDict = New Dictionary(Of TValue, TValue)(additionalValues.Length + 1)
            Else
                myDict = New Dictionary(Of TValue, TValue)(additionalValues.Length + 1, valueComparer)
            End If
            If (value IsNot Nothing) Then
                myDict.Add(value, value)
            End If
            For Each myValue As TValue In additionalValues
                If (myValue Is Nothing) Then Continue For
                myDict.Item(myValue) = myValue
            Next
            'Return the combination
            Return GetInstanceByValuesOptimized(myDict)
        End Function

        ''' <summary>Returns a combination where all possible members are set.</summary>
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Public Shared ReadOnly Property All() As TCombi
            Get
                Dim myResult As TCombi = _All
                If (myResult Is Nothing) Then
                    'Define the All combination
                    Dim myMembers As TEnum() = Members
                    Select Case myMembers.Length
                        Case 0
                            myResult = Empty
                        Case 1
                            myResult = myMembers(0).ToCombi()
                        Case Else
                            myResult = New TCombi()
                            myResult.SetAll()
                    End Select
                    'Assign the result
                    _All = myResult
                End If
                'Return the result
                Return myResult
            End Get
        End Property

        ''' <summary>Returns an empty combination.</summary>
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        <SuppressMessage("Microsoft.Design", "CA1000:DoNotDeclareStaticMembersOnGenericTypes", Justification:="It is needed and cannot be changed into an instance method.")> _
        Public Shared ReadOnly Property Empty() As TCombi
            Get
                Dim myResult As TCombi = _Empty
                If (myResult Is Nothing) Then
                    myResult = New TCombi()
                    _Empty = myResult
                End If
                Return myResult
            End Get
        End Property

        'Public Properties

        ''' <summary>Gets the member of this combination with the given index. If the index is out-of-range, null is returned.</summary>
        ''' <param name="index">The zero-based index of the member to get.</param>
        Default Public ReadOnly Property Item(ByVal index As Int32) As TEnum
            Get
                'Handle out-of-range
                If (index < 0) Then Return Nothing
                If (index >= Count) Then Return Nothing
                'Find according flag
                Dim myFlags As Boolean() = Flags
                For i As Int32 = 0 To myFlags.Length - 1
                    If (myFlags(i)) Then
                        If (index = 0) Then Return Members(i)
                        index -= 1
                    End If
                Next
                'Should never reach here
                Throw New InvalidOperationException("Internal error in " & Me.GetType().FullName & "!")
            End Get
        End Property

        ''' <summary>Returns the number of members this combination has set. This property is thread-safe.</summary>
        Public ReadOnly Property Count As Int32
            Get
                'Return result from cache
                Dim myResult As Int32 = _Count
                If (myResult > -1) Then Return myResult
                'Initialize result
                Dim myFlags As Boolean() = Flags
                myResult = 0
                For Each myFlag As Boolean In myFlags
                    If (myFlag) Then myResult += 1
                Next
                'Assign result
                _Count = myResult
                'Return result
                Return myResult
            End Get
        End Property

        ''' <summary>Returns true if none of the flags is set, false otherwise. This property is thread-safe.</summary>
        Public ReadOnly Property IsEmpty() As Boolean
            Get
                Dim myCount As Int32 = Count
                Return (myCount = 0)
            End Get
        End Property

        ''' <summary>Returns true if all of the flags are set, false otherwise. If the enum does not define
        ''' any members, true is returned. This property is thread-safe.</summary>
        Public ReadOnly Property IsAllSet() As Boolean
            Get
                'Return true is Count is equal to the number of flags
                Dim myFlags As Boolean() = Flags
                Dim myCount As Int32 = Count
                Return (myFlags.Length = myCount)
            End Get
        End Property

        'Public Methods

        ''' <summary>Determines whether this instance equals another instance (it may be equal to a member or to a combination).</summary>
        ''' <param name="obj">The other instance that is supposed to be equal.</param>
        <EditorBrowsable(EditorBrowsableState.Advanced)> _
        Public Overrides Function Equals(ByVal obj As Object) As Boolean
            'Handle null
            If (obj Is Nothing) Then Return False
            'Handle same reference
            If (Object.ReferenceEquals(Me, obj)) Then Return True
            'Handle Combination
            Dim myCombination As Combi = TryCast(obj, Combi)
            If (myCombination IsNot Nothing) Then
                Return Equals(myCombination)
            End If
            'Handle Member
            Dim myMember As TEnum = TryCast(obj, TEnum)
            If (myMember IsNot Nothing) Then
                Return Equals(myMember)
            End If
            'Handle Value
            If (TypeOf obj Is TValue) Then
                Dim myValue As TValue = DirectCast(obj, TValue)
                Return Equals(myValue)
            End If
            'Otherwise return false
            Return False
        End Function

        ''' <summary>Compares this combination with another one and returns true if they are equal (the same flags are set),
        ''' false otherwise.</summary>
        ''' <param name="other">The members that are supposed to be equal.</param>
        ''' <returns>True if exactly the same members are set, false otherwise.</returns>
        Public Overloads Function Equals(ByVal other As Combi) As Boolean Implements IEquatable(Of Combi).Equals
            'Handle same reference
            If (Object.ReferenceEquals(Me, other)) Then Return True
            'Handle null
            If (other Is Nothing) Then Return False
            'Compare count first
            If (Me.Count <> other.Count) Then Return False
            'Compare the flags
            Dim myFlagsX As Boolean() = Me.Flags
            Dim myFlagsY As Boolean() = other.Flags
            For i As Int32 = 0 To myFlagsX.Length - 1
                If (myFlagsX(i) <> myFlagsY(i)) Then Return False
            Next
            'Everything is the same, return true
            Return True
        End Function

        ''' <summary>Compares this combination with the given member and returns true if this combination consists only of
        ''' that member, false otherwise.</summary>
        ''' <param name="other">The member this combination is compared with</param>
        Public Overloads Function Equals(ByVal other As TEnum) As Boolean Implements IEquatable(Of TEnum).Equals
            If (other Is Nothing) Then Return False
            If (Count <> 1) Then Return False
            Dim myMember As TEnum = Item(0)
            Return (myMember.Equals(other))
        End Function

        ''' <summary>Compares all members of this combination with the given value and returns true if they all match 
        ''' the value, false otherwise. If the given value is null, false is returned. If the combination is empty,
        ''' false is returned.</summary>
        ''' <param name="other">The value this combination is compared with</param>
        Public Overloads Function Equals(ByVal other As TValue) As Boolean
            If (other Is Nothing) Then Return False
            Dim myMembers As TEnum() = ToArray()
            If (myMembers.Length = 0) Then Return False
            For Each myMember As TEnum In myMembers
                If (myMember.Equals(other)) Then Continue For
                Return False
            Next
            'Otherwise return true
            Return True
        End Function

        ''' <summary>Retrieves an XOR combination of all set members. This ensures that a member and a Combination that contains 
        ''' only one member have the same hashcode.</summary>
        <EditorBrowsable(EditorBrowsableState.Advanced)> _
        Public Overrides Function GetHashCode() As Int32
            Dim myResult As Int32? = _HashCode
            If (myResult Is Nothing) Then
                'Calculate hashcode
                Dim myHashCode As Int32 = 0
                Dim myMembers As TEnum() = ToArray()
                For Each myMember As TEnum In myMembers
                    myHashCode = myHashCode Xor myMember.GetHashCode()
                Next
                'Assign and return
                _HashCode = myHashCode
                Return myHashCode
            End If
            'Return from cache
            Return myResult.Value
        End Function

        ''' <summary>Returns true if the given member is set, false otherwise (null also returns false).</summary>
        ''' <param name="member">The member to check</param>
        Public Function Contains(ByVal member As TEnum) As Boolean
            'Ignore if null
            If (member Is Nothing) Then Return False
            'Return whether the flag is set
            Return Flags(member.Index)
        End Function

        ''' <summary>Returns true if all given members are set, false otherwise (if the combination does not contain any members, 
        ''' false is returned).</summary>
        ''' <param name="members">The members to check</param>
        Public Function Contains(ByVal members As TCombi) As Boolean
            'Ignore if null
            If (members Is Nothing) OrElse (members.IsEmpty) Then Return False
            If (members.Count > Me.Count) Then Return False
            'Compare the flags
            Dim myFlags As Boolean() = Me.Flags
            Dim myMembers As TEnum() = members.ToArray()
            For Each myMember As TEnum In myMembers
                If (Not myFlags(myMember.Index)) Then Return False
            Next
            Return True
        End Function

        ''' <summary>Returns true if all given members are set, false otherwise (null values are ignored; if the array does not contain
        ''' any members, false is returned).</summary>
        ''' <param name="member">The members to check</param>
        Public Function Contains(ByVal ParamArray member As TEnum()) As Boolean
            'Ignore if null
            If (member Is Nothing) Then Return False
            'Set the according flags to false
            Dim myHasFlags As Boolean = False
            Dim myFlags As Boolean() = Me.Flags
            For Each myMember As TEnum In member
                If (myMember Is Nothing) Then Continue For
                If (Not myFlags(myMember.Index)) Then Return False
                myHasFlags = True
            Next
            'Return true if at least one flag was given
            Return myHasFlags
        End Function

        ''' <summary>Returns true if at least one of the given members is set, false otherwise. This operation is thread-safe. If 
        ''' members is null, false is returned.</summary>
        ''' <param name="members">The members to check</param>
        Public Function ContainsAny(ByVal members As TCombi) As Boolean
            'Ignore if null
            If (members Is Nothing) OrElse (members.IsEmpty) OrElse (Me.IsEmpty) Then Return False
            'Compare the flags
            Dim myFlags As Boolean() = Me.Flags
            Dim myMembers As TEnum() = members.ToArray()
            For Each myMember As TEnum In myMembers
                'Return true if set
                If (myFlags(myMember.Index)) Then Return True
            Next
            'Otherwise return false
            Return False
        End Function

        ''' <summary>Returns true if at least one of the given members is set, false otherwise. This operation is not thread-safe, 
        ''' please ensure the members collection is not manipulated during the time this method executes. If the collection is
        ''' null or empty, false is returned.</summary>
        ''' <param name="members">The members to check</param>
        Public Function ContainsAny(members As IEnumerable(Of TEnum)) As Boolean
            'Convert to array
            If (members Is Nothing) OrElse (Me.IsEmpty) Then Return False
            'Compare the flags
            Dim myFlags As Boolean() = Me.Flags
            For Each myMember As TEnum In members
                'Return true if set
                If (myMember Is Nothing) Then Continue For
                If (myFlags(myMember.Index)) Then Return True
            Next
            'Otherwise return false
            Return False
        End Function

        ''' <summary>Returns true if at least one of the given members is set, false otherwise. This operation is thread-safe. 
        ''' Null values are ignored.</summary>
        ''' <param name="member">The first member to check.</param>
        ''' <param name="additionalMembers">Additional members to check (split into two parameters for easier consumation by C#).</param>
        Public Function ContainsAny(member As TEnum, ByVal ParamArray additionalMembers As TEnum()) As Boolean
            'Handle empty array
            If (Me.IsEmpty) Then Return False
            Dim myFlags As Boolean() = Me.Flags
            If (additionalMembers Is Nothing) OrElse (additionalMembers.Length = 0) Then
                If (member Is Nothing) Then Return False
                Return myFlags(member.Index)
            End If
            'Compare the flags
            If (member IsNot Nothing) AndAlso (myFlags(member.Index)) Then Return True
            For Each myMember As TEnum In additionalMembers
                'Return true if set
                If (myMember Is Nothing) Then Continue For
                If (myFlags(myMember.Index)) Then Return True
            Next
            'Otherwise return false
            Return False
        End Function

        ''' <summary>Returns the set members as an array (.NET 2.0 support).</summary>
        Public Function ToArray() As TEnum()
            'Hint: May not be cached (or only for Array.Copy())
            Dim myResult As New List(Of TEnum)
            Dim myFlags As Boolean() = Flags
            Dim myMembers As TEnum() = Members
            For i As Int32 = 0 To myFlags.Length - 1
                If (myFlags(i)) Then myResult.Add(myMembers(i))
            Next
            Return myResult.ToArray()
        End Function

        Public Overrides Function ToString() As String
            'Write all set members
            Dim myResult As New StringBuilder()
            Dim myMembers As TEnum() = ToArray()
            If (myMembers.Length = 0) Then Return ""
            If (myMembers.Length = 1) Then Return myMembers(0).Name
            For Each myMember As TEnum In ToArray()
                myResult.Append(myMember.Name)
                myResult.Append(", ")
            Next
            myResult.Length -= 2
            Return myResult.ToString()
        End Function

        'Public Operators 

        ''' <summary>Every member is castable into a Combination. To avoid ambiguity for the operators this operation is
        ''' defined as explicite cast even if it cannot fail. This operation is thread-safe.</summary>
        ''' <param name="member">The member to be cast into a Combination.</param>
        Public Shared Narrowing Operator CType(ByVal member As TEnum) As Combi
            If (member Is Nothing) Then Return Combi.Empty
            Return member.ToCombi()
        End Operator

        ''' <summary>If the combination consists of exactly one member, it can be cast into the member. If no member or more than one
        ''' member is set, an InvalidCastException is thrown. This operation is thread-safe.</summary>
        ''' <param name="combination">The combination containing a single member to be cast.</param>
        ''' <exception cref="InvalidCastException">An InvalidCastException is thrown if the combination has less/more than one member set.</exception>
        Public Shared Narrowing Operator CType(ByVal combination As Combi) As TEnum
            'Check args
            If (combination Is Nothing) OrElse (combination.Count = 0) Then
                Throw New InvalidCastException("The combination has less than one member set.")
            End If
            If (combination.Count > 1) Then Throw New InvalidCastException("The combination has more than one member set.")
            'Return the result
            Dim myMembers As TEnum() = combination.ToArray()
            Return myMembers(0)
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. It is more efficient to use the construtor that takes a paramarray to combine more than two members.
        ''' This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a member.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Or(ByVal arg1 As Combi, ByVal arg2 As TEnum) As TCombi
            Return (arg2 Or arg1)
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a member.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Or(ByVal arg1 As TEnum, ByVal arg2 As Combi) As TCombi
            'Handle null
            If (arg1 Is Nothing) Then
                If (arg2 Is Nothing) Then Return Combi.Empty
                Return CType(arg2, TCombi)
            End If
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return arg1.ToCombi()
            'Return arg2 if the flag was already set
            If (arg2.Contains(arg1)) Then Return CType(arg2, TCombi)
            'Merge both flags
            Dim myResult As TCombi = Combi.GetInstanceOptimized(arg1, arg2.ToArray())
            Return myResult
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. If one of the arguments is null or empty, the other argument is returned (same instance). If both 
        ''' arguments are null, <see cref="Combi.Empty">Combi.Empty</see> is returned. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members (null is allowed).</param>
        ''' <param name="arg2">The second argument, a combination of members (null is allowed).</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Or(ByVal arg1 As Combi, ByVal arg2 As TCombi) As TCombi
            'Handle null
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) Then
                If (arg2 Is Nothing) Then Return Combi.Empty
                Return arg2
            End If
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return CType(arg1, TCombi) 'same instance
            'Handle equal
            If (arg1 = arg2) Then Return arg2
            'Handle contains
            Dim myCount1 As Int32 = arg1.Count
            Dim myCount2 As Int32 = arg2.Count
            If (myCount1 > myCount2) Then
                If (arg1.Contains(arg2)) Then
                    Return CType(arg1, TCombi)
                End If
            ElseIf (myCount1 < myCount2) Then
                If (arg2.Contains(CType(arg1, TCombi))) Then
                    Return arg2
                End If
            End If
            'Merge both flags
            Dim myResult As TCombi = Combi.GetInstanceRaw()
            Dim myFlags As Boolean() = myResult.Flags
            Dim myFlagsX As Boolean() = arg1.Flags
            Dim myFlagsY As Boolean() = arg2.Flags
            For i As Int32 = 0 To myFlags.Length - 1
                myFlags(i) = (myFlagsX(i) OrElse myFlagsY(i))
            Next
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination using a binary XOR operation. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a member.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Xor(ByVal arg1 As Combi, ByVal arg2 As TEnum) As TCombi
            'Handle null/empty
            If (arg2 Is Nothing) Then
                If (arg1 Is Nothing) Then Return Combi.Empty
                Return CType(arg1, TCombi)
            End If
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) Then Return arg2.ToCombi()
            'Merge both flags
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg1)
            myResult.Toggle(arg2)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination using a binary XOR operation. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a member.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Xor(ByVal arg1 As TEnum, ByVal arg2 As Combi) As TCombi
            'Handle null/empty
            If (arg1 Is Nothing) Then
                If (arg2 Is Nothing) Then Return Combi.Empty
                Return CType(arg2, TCombi)
            End If
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return arg1.ToCombi()
            'Merge both flags
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg2)
            myResult.Toggle(arg1)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination using a binary XOR operation. To toggle all
        ''' members use an unary NOT operation. If one of the arguments is null or empty, the other argument is returned 
        ''' (same instance). If both arguments are null, an empty combination is returned. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Xor(ByVal arg1 As Combi, ByVal arg2 As TCombi) As TCombi
            'Handle null/empty
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) Then
                If (arg2 Is Nothing) Then
                    Return Combi.Empty
                End If
                Return arg2 'same instance
            End If
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return CType(arg1, TCombi) 'same instance
            'Merge both flags
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg1)
            myResult.Toggle(arg2)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. It is more efficient to use the construtor that takes a paramarray to combine more than two members. 
        ''' This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a member.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator +(ByVal arg1 As Combi, ByVal arg2 As TEnum) As TCombi
            Return (arg1 Or arg2)
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a member.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator +(ByVal arg1 As TEnum, ByVal arg2 As Combi) As TCombi
            Return (arg1 Or arg2)
        End Operator

        ''' <summary>Combines both arguments into a new instance of Combination. There is no difference between the "OR" and the 
        ''' "+" operator. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator +(ByVal arg1 As Combi, ByVal arg2 As TCombi) As TCombi
            Return (arg1 Or arg2)
        End Operator

        ''' <summary>Returns a copy of arg1 (a new instance) without member arg2. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The argument to subtract, a member.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator -(ByVal arg1 As Combi, ByVal arg2 As TEnum) As TCombi
            'Handle null
            If (arg1 Is Nothing) Then Return Combi.Empty
            If (arg2 Is Nothing) Then Return CType(arg1, TCombi)
            'Return arg1 if the flag is not set
            If (Not arg1.Contains(arg2)) Then
                Return CType(arg1, TCombi)
            End If
            'Otherwise return new instance
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg1)
            myResult.Remove(arg2)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Returns a copy of arg1 (a new instance) without the members of arg2. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The argument to subtract, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator -(ByVal arg1 As Combi, ByVal arg2 As TCombi) As TCombi
            'Handle null
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) Then Return Combi.Empty
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return CType(arg1, TCombi)
            'Return same instance if there was no match
            If (Not arg1.ContainsAny(arg2)) Then
                Return CType(arg1, TCombi)
            End If
            'Remove the flags
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg1)
            myResult.Remove(arg2)
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Returns a new instance of Combination that contains all members that are in both arg1 and arg2. This operation is
        ''' thread-safe.</summary>
        ''' <param name="arg1">The first argument, a combination of members.</param>
        ''' <param name="arg2">The second argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator And(ByVal arg1 As Combi, ByVal arg2 As TCombi) As TCombi
            'Handle null
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) OrElse (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return Combi.Empty
            'Handle no match
            If (Not arg1.ContainsAny(arg2)) Then
                Return Combi.Empty
            End If
            'Merge both flags
            Dim myResult As New TCombi()
            Dim myFlags As Boolean() = myResult.Flags
            Dim myFlagsX As Boolean() = arg1.Flags
            Dim myFlagsY As Boolean() = arg2.Flags
            For i As Int32 = 0 To myFlags.Length - 1
                myFlags(i) = (myFlagsX(i) AndAlso myFlagsY(i))
            Next
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Returns a new instance of Combination that contains all members that are not in arg. This 
        ''' operation is thread-safe.</summary>
        ''' <param name="arg">The argument, a combination of members.</param>
        <SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates", Justification:="As there is always a new instance returned it would be rather confusing to have an instance method.")> _
        Public Shared Operator Not(ByVal arg As Combi) As TCombi
            'Handle well-known collections
            If (arg Is Nothing) OrElse (arg.IsEmpty) Then Return Combi.All
            If (arg.IsAllSet) Then Return Combi.Empty
            'Return new inverted collection
            Dim myResult As TCombi = Combi.GetInstanceRaw(arg)
            myResult.ToggleAll()
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Operator

        ''' <summary>Returns true if the given combination has only the given member set, false otherwise. If one or
        ''' both of the arguments are null, false is returned. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a combination of members.</param>
        ''' <param name="arg2">The second argument to compare, a member.</param>
        Public Shared Operator =(ByVal arg1 As Combi, ByVal arg2 As TEnum) As Boolean
            If (arg1 Is Nothing) Then Return False
            Return arg1.Equals(arg2)
        End Operator

        ''' <summary>Returns true if the given combination has only the given member set, false otherwise. If one or
        ''' both of the arguments are null, false is returned. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a member.</param>
        ''' <param name="arg2">The second argument to compare, a combination of members.</param>
        Public Shared Operator =(ByVal arg1 As TEnum, ByVal arg2 As Combi) As Boolean
            If (arg2 Is Nothing) Then Return False
            Return arg2.Equals(arg1)
        End Operator

        ''' <summary>Returns true if the given combination has only one member set whos value equals the given one, false otherwise. 
        ''' This operation is thread-safe if the value instance and the value's comparer are thread-safe (or not accessed by other
        ''' threads).</summary>
        ''' <param name="arg1">The first argument to compare, a combination of members.</param>
        ''' <param name="arg2">The second argument to compare, a member.</param>
        Public Shared Operator =(ByVal arg1 As Combi, ByVal arg2 As TValue) As Boolean
            If (arg1 Is Nothing) Then Return False
            Return arg1.Equals(arg2)
        End Operator

        ''' <summary>Returns true if the given combination has only one member set whos value equals the given one, false otherwise. 
        ''' This operation is thread-safe if the value instance and the value's comparer are thread-safe (or not accessed by other
        ''' threads).</summary>
        ''' <param name="arg1">The first argument to compare, a member.</param>
        ''' <param name="arg2">The second argument to compare, a combination of members.</param>
        Public Shared Operator =(ByVal arg1 As TValue, ByVal arg2 As Combi) As Boolean
            If (arg2 Is Nothing) Then Return False
            Return arg2.Equals(arg1)
        End Operator

        ''' <summary>Returns true if the given combinations have the same members. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a combination of members.</param>
        ''' <param name="arg2">The second argument to compare, a combination of members.</param>
        Public Shared Operator =(ByVal arg1 As Combi, ByVal arg2 As TCombi) As Boolean
            'Handle null
            If (arg1 Is Nothing) OrElse (arg1.IsEmpty) Then
                If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return True
                Return False
            End If
            If (arg2 Is Nothing) OrElse (arg2.IsEmpty) Then Return False
            'Compare the values
            Return arg1.Equals(arg2)
        End Operator

        ''' <summary>Returns false if the given combination has only the given member set, true otherwise. (If the member is null
        ''' and also the combination is null or empty, false is returned to be consistent to the explicite cast operator). This 
        ''' operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a combination of members.</param>
        ''' <param name="arg2">The second argument to compare, a member.</param>
        Public Shared Operator <>(ByVal arg1 As Combi, ByVal arg2 As TEnum) As Boolean
            Return (Not (arg1 = arg2))
        End Operator

        ''' <summary>Returns false if the given combination has only the given member set, true otherwise. (If the member is null
        ''' and also the combination is null or empty, false is returned to be consistent to the explicite cast operator). This 
        ''' operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a member.</param>
        ''' <param name="arg2">The second argument to compare, a combination of members.</param>
        Public Shared Operator <>(ByVal arg1 As TEnum, ByVal arg2 As Combi) As Boolean
            Return (Not (arg1 = arg2))
        End Operator

        ''' <summary>Returns false if the given combination has only one member set whos value equals the given one, true otherwise. 
        ''' This operation is thread-safe if the value instance and the value's comparer are thread-safe (or not accessed by other
        ''' threads).</summary>
        ''' <param name="arg1">The first argument to compare, a combination of members.</param>
        ''' <param name="arg2">The second argument to compare, a member.</param>
        Public Shared Operator <>(ByVal arg1 As Combi, ByVal arg2 As TValue) As Boolean
            Return (Not (arg1 = arg2))
        End Operator

        ''' <summary>Returns false if the given combination has only one member set whos value equals the given one, true otherwise. 
        ''' This operation is thread-safe if the value instance and the value's comparer are thread-safe (or not accessed by other
        ''' threads).</summary>
        ''' <param name="arg1">The first argument to compare, a member.</param>
        ''' <param name="arg2">The second argument to compare, a combination of members.</param>
        Public Shared Operator <>(ByVal arg1 As TValue, ByVal arg2 As Combi) As Boolean
            Return (Not (arg1 = arg2))
        End Operator

        ''' <summary>Returns false if the given combinations have the same members. If both combinations are null or empty (in any 
        ''' combination), false is returned. This operation is thread-safe.</summary>
        ''' <param name="arg1">The first argument to compare, a combination of members.</param>
        ''' <param name="arg2">The second argument to compare, a combination of members.</param>
        Public Shared Operator <>(ByVal arg1 As Combi, ByVal arg2 As TCombi) As Boolean
            Return (Not (arg1 = arg2))
        End Operator

        'Private Properties

        ''' <summary>Initializes and returns the flags.</summary>
        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private ReadOnly Property Flags As Boolean()
            Get
                Dim myResult As Boolean() = _Flags
                If (myResult Is Nothing) Then
                    myResult = New Boolean(ValueEnum(Of TEnum, TValue, TCombi).Members.Length - 1) {}
                    _Flags = myResult
                End If
                Return myResult
            End Get
        End Property

        ''' <summary>For debugging only, visualization for property Flags.</summary>
        Private ReadOnly Property AllMembers() As MemberFlag()
            Get
                Dim myMembers As TEnum() = GetMembers()
                Dim myFlags As Boolean() = Flags
                Dim myResult(myMembers.Length - 1) As MemberFlag
                For i As Int32 = 0 To myMembers.Length - 1
                    myResult(i) = New MemberFlag(myMembers(i), myFlags(i))
                Next
                Return myResult
            End Get
        End Property

        ''' <summary>For debugging only, calls function ToArray().</summary>
        Private ReadOnly Property AssignedMembers() As MemberValue()
            Get
                Dim myMembers As TEnum() = ToArray()
                Dim myResult(myMembers.Length - 1) As MemberValue
                For i As Int32 = 0 To myMembers.Length - 1
                    myResult(i) = New MemberValue(myMembers(i))
                Next
                Return myResult
            End Get
        End Property

        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private ReadOnly Property DebuggerDisplayValue As String
            Get
                'Write all set members
                Dim myResult As New StringBuilder()
                Dim myMembers As TEnum() = ToArray()
                If (myMembers.Length = 0) Then Return "[empty]"
                If (myMembers.Length = 1) Then Return myMembers(0).Name
                For Each myMember As TEnum In ToArray()
                    myResult.Append(myMember.Name)
                    myResult.Append(", ")
                Next
                myResult.Length -= 2
                Return myResult.ToString()
            End Get
        End Property

        'Private Methods

        Private Shared Function GetInstanceByValuesOptimized(ByVal valueDict As Dictionary(Of TValue, TValue)) As TCombi
            'Determine result
            Dim myResult As New TCombi()
            Dim myFlags As Boolean() = myResult.Flags
            Dim myPossibleMembers As TEnum() = Members()
            For i As Int32 = 0 To myPossibleMembers.Length - 1
                Dim myMember As TEnum = myPossibleMembers(i)
                If (valueDict.ContainsKey(myMember._Value)) Then
                    myFlags(i) = True
                End If
            Next
            'Optimize result
            myResult = OptimizeResult(myResult)
            'Return result
            Return myResult
        End Function

        ''' <summary>Why optimize the result? Because 1) comparing same instances is much faster than comparing instances
        ''' that are only equal, 2) garbage collection of level 0 objects is rather cheap, 3) we use a little less memory
        ''' and 4) it is does not cost much to perform these optimizations (the initialization of Count is the only somewhat
        ''' costy operation).</summary>
        ''' <param name="result">A combination containing the same members.</param>
        Private Shared Function OptimizeResult(ByVal result As TCombi) As TCombi
            If (result Is Nothing) Then Return Combi.Empty
            Select Case result.Count
                Case 0
                    Return Combi.Empty
                Case 1
                    Return result(0).ToCombi()
                Case Else
                    If (result.IsAllSet) Then Return Combi.All
            End Select
            Return result
        End Function

        ''' <summary>Sets the given member (if the member is already set or <paramref name="member" /> is null, it is ignored). 
        ''' An <see cref="InvalidOperationException" /> is thrown if the instance is protected. The Set method is similair to a 
        ''' binary OR operation.</summary>
        ''' <param name="member">The member to set.</param>
        Private Sub [Set](ByVal member As TEnum)
            'Ignore if null
            If (member Is Nothing) Then Return
            'Set the flag
            Flags(member.Index) = True
        End Sub

        ''' <summary>Sets the given members (null values and duplicates are ignored). An <see cref="InvalidOperationException" /> 
        ''' is thrown if the instance is protected. The Set method is similair to a binary OR operation.</summary>
        ''' <param name="member">The members to set.</param>
        Private Sub [Set](ByVal member As TEnum())
            'Ignore if null
            If (member Is Nothing) Then Return
            'Set the according flags
            Dim myFlags As Boolean() = Me.Flags
            For Each myMember As TEnum In member
                If (myMember Is Nothing) Then Continue For
                myFlags(myMember.Index) = True
            Next
        End Sub

        ''' <summary>Sets all members. An <see cref="InvalidOperationException" /> is thrown if the instance is protected.</summary>
        Private Sub SetAll()
            'Set all flags to true
            Dim myFlags As Boolean() = Flags
            For i As Int32 = 0 To myFlags.Length - 1
                myFlags(i) = True
            Next
        End Sub

        ''' <summary>Toggles a member (if it is set, it is removed; if it is not set, it is set). If the member is null, it is ignored. 
        ''' An <see cref="InvalidOperationException" /> is thrown if the instance is protected. Toggle is similair to a binary 
        ''' XOR operation.</summary>
        ''' <param name="member">The member to toggle.</param>
        Private Sub Toggle(ByVal member As TEnum)
            'Ignore if null
            If (member Is Nothing) Then Return
            'Toggle the flag
            Dim myFlags As Boolean() = Flags
            Dim myIndex As Int32 = member.Index
            myFlags(myIndex) = (Not myFlags(myIndex))
        End Sub

        ''' <summary>Merges the given combination using a binary XOR operation.</summary>
        ''' <param name="combination">The members to toggle.</param>
        Private Sub Toggle(ByVal combination As Combi)
            'Handle null
            If (combination Is Nothing) OrElse (combination.IsEmpty) Then Return
            'Merge both flags
            Dim myFlagsX As Boolean() = Flags
            Dim myFlagsY As Boolean() = combination.Flags
            For i As Int32 = 0 To myFlagsX.Length - 1
                Dim myX As Boolean = myFlagsX(i)
                Dim myY As Boolean = myFlagsY(i)
                If (myX AndAlso myY) Then
                    myFlagsX(i) = False
                ElseIf (myX OrElse myY) Then
                    myFlagsX(i) = True
                End If
            Next
        End Sub

        ''' <summary>Toggles all members. An <see cref="InvalidOperationException" /> is thrown if the instance is protected. ToggleAll is 
        ''' similair to an unary NOT operation.</summary>
        Private Sub ToggleAll()
            'Toggle all flags
            Dim myFlags As Boolean() = Flags
            For i As Int32 = 0 To myFlags.Length - 1
                myFlags(i) = (Not myFlags(i))
            Next
        End Sub

        ''' <summary>Removes the given member (if the member is not set or <paramref name="member" /> is null, it is ignored).
        ''' An <see cref="InvalidOperationException" /> is thrown if the instance is protected. The Remove method is similair to 
        ''' a binary NOT operation.</summary>
        ''' <param name="member">The member to remove</param>
        Private Sub Remove(ByVal member As TEnum)
            'Ignore if null
            If (member Is Nothing) Then Return
            'Remove the flag
            Flags(member.Index) = False
        End Sub

        ''' <summary>Removes the given members (null values and duplicates are ignored). An <see cref="InvalidOperationException" /> is 
        ''' thrown if the instance is protected. The Remove method is similair to a binary NOT operation (does not exist in Visual Basic).</summary>
        ''' <param name="combination">The members to remove</param>
        Private Sub Remove(ByVal combination As Combi)
            'Ignore if null
            If (combination Is Nothing) OrElse (combination.IsEmpty) Then Return
            'Set the according flags to false
            Dim myFlags As Boolean() = Me.Flags
            Dim myOtherMembers As TEnum() = combination.ToArray()
            For Each myMember As TEnum In myOtherMembers
                myFlags(myMember.Index) = False
            Next
        End Sub

        ''' <summary>Retrieves an enumerator for the set members. This method takes a snapshot of the members before returning the 
        ''' enumerator, allowing you to manipulate the members even from within the loop.</summary>
        Private Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of TEnum) Implements System.Collections.Generic.IEnumerable(Of TEnum).GetEnumerator
            Dim myResult As IEnumerable(Of TEnum) = ToArray()
            Return myResult.GetEnumerator()
        End Function

        ''' <summary>Retrieves an enumerator for the set members. This method takes a snapshot of the members before returning the 
        ''' enumerator, allowing you to manipulate the members even from within the loop.</summary>
        Private Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function


        '**********************************************************************
        ' INNER CLASS: MemberFlag
        '**********************************************************************

        ''' <summary>For debugged display.</summary>
        <DebuggerDisplay("{Flag}", Name:="{Name}")> _
        Private Class MemberFlag

            'Public Fields

            <SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields", Justification:="It is used to be displayed in the debugger.")> _
            <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
            Public ReadOnly Name As String
            <SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields", Justification:="It is used to be displayed in the debugger.")> _
            <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
            Public ReadOnly Flag As Boolean
            <SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields", Justification:="It is used to be displayed in the debugger.")> _
            Public ReadOnly Member As TEnum

            'Constructors

            Public Sub New(ByVal member As TEnum, ByVal flag As Boolean)
                Me.Name = member.Name
                Me.Flag = flag
                Me.Member = member
            End Sub

        End Class


        '**********************************************************************
        ' INNER CLASS: MemberValue
        '**********************************************************************

        ''' <summary>For debugged display.</summary>
        <DebuggerDisplay("{Value}", Name:="{Name}")> _
        Private Class MemberValue

            'Public Fields

            <SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields", Justification:="It is used to be displayed in the debugger.")> _
            <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
            Public ReadOnly Name As String
            <SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields", Justification:="It is used to be displayed in the debugger.")> _
            <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
            Public ReadOnly Value As TValue
            <SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields", Justification:="It is used to be displayed in the debugger.")> _
            Public ReadOnly Member As TEnum

            'Constructors

            Public Sub New(ByVal member As TEnum)
                Me.Name = member.Name
                Me.Value = member.Value
                Me.Member = member
            End Sub

        End Class

    End Class

End Class



''' <summary>Base class of valued custom enums that do not need to overwrite the combination class.</summary>
''' <typeparam name="TEnum">The type of the enum.</typeparam>
''' <typeparam name="TValue">The type of the value.</typeparam>
<SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix", Justification:="Because it is a base class for all enums, the suffix ""Enum"" is just fine.")> _
Public MustInherit Class ValueEnum(Of TEnum As ValueEnum(Of TEnum, TValue), TValue)
    Inherits ValueEnum(Of TEnum, TValue, Combi)

    'Constructors

    ''' <summary>Called by implementors to create a new instance of TEnum (when assigning the instance to a static field). 
    ''' Important: Make your constructors private to ensure there are no instances except the ones initialized 
    ''' by your subclass! Null values are not supported and throw an ArgumentNullException.</summary>
    ''' <param name="value">This member's value (not null)</param>
    ''' <exception cref="ArgumentNullException">An ArgumentNullException is thrown if aValue is null.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this instance's type is not of type <typeparam name="TEnum" />.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException may be thrown (check only in full-trust possible) if this instance has non-private constructors.</exception>
    Protected Sub New(ByVal value As TValue)
        MyBase.New(value)
    End Sub

    ''' <summary>Called by implementors to create a new instance of TEnum (when assigning the instance to a static field). 
    ''' Important: Make your constructors private to ensure there are no instances except the ones initialized 
    ''' by your subclass! Null values are not supported and throw an ArgumentNullException.</summary>
    ''' <param name="value">This member's value (not null)</param>
    ''' <param name="caseSensitive">Leave null for automatic determination (recommended), or set explicitely.</param>
    ''' <param name="valueComparer">What comparer should be used to compare the values in method <see cref="GetMemberByValue" /> and <see cref="Equals" /> as well as the equal operator or null to use Object.Equals(..).</param>
    ''' <exception cref="ArgumentNullException">An ArgumentNullException is thrown if aValue is null.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException is thrown if this instance's type is not of type <typeparam name="TEnum" />.</exception>
    ''' <exception cref="InvalidOperationException">An InvalidOperationException may be thrown (check only in full-trust possible) if this instance has non-private constructors.</exception>
    Protected Sub New(ByVal value As TValue, ByVal caseSensitive As Boolean?, ByVal valueComparer As IEqualityComparer(Of TValue))
        MyBase.New(value, caseSensitive, valueComparer)
    End Sub


    '**********************************************************************
    ' INNER CLASS: Combi
    '**********************************************************************

    <SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId:="Combi", Justification:="""Combination"" was too long.")> _
    <SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix", Justification:="Trust me, Combi is just fine.")> _
    <SuppressMessage("Microsoft.Design", "CA1034:NestedTypesShouldNotBeVisible")> _
    <EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Shadows Class Combi
        Inherits ValueEnum(Of TEnum, TValue, Combi).Combi

    End Class

End Class


''' <summary>The CustomEumIgnoreAttribute is used to flag public static readonly fields and -properties 
''' not to be treated as members (not to be included in Combi.All, GetMembers(), duplicate checks etc.) 
''' but as simple fields/properties that happen to have the same signature. This allows for assigning the same 
''' member instance to additional fields like a default field that takes one of the other members as value.
''' </summary>
<AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property)> _
Public NotInheritable Class CustomEnumIgnoreAttribute
    Inherits Attribute

End Class
