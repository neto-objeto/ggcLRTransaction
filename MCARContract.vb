
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     MCARContract Object
'
' Copyright 2020 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  TEEJEI [ 01/17/2026 09:50:59 ]
'       Started creating this object.
'  
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Strict Off

'==================== IMPORTS ====================
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports rmjGOCAS
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports Microsoft.VisualBasic.ApplicationServices

'==================== CLASS ====================
Public Class MCARContract

    '==================== PRIVATE VARIABLES ====================
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode

    Private p_sBranchCd As String
    Private p_sBranchNm As String
    Private p_nTranStat As Int32
    Private p_sParent As String

    Private p_oClientName As ggcClient.Client
    Private psMessage As String
    Private p_sCSSNumbr As String

    '==================== CONSTANTS ====================
    Private Const p_sMasTable As String = "MC_AR_Contract_Info"
    Private Const p_sMsgHeadr As String = "MC_AR_Contract_Info"
    Private Const pxeMaxReferx As Integer = 3

    '==================== CONSTRUCTORS ====================
    Public Sub New(ByVal foRider As GRider, ByVal fnStatus As Int32)
        Me.New(foRider)
        p_nTranStat = fnStatus
    End Sub

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_nEditMode = xeEditMode.MODE_UNKNOWN

        p_sBranchCd = p_oApp.BranchCode
        p_sBranchNm = p_oApp.BranchName
        p_nTranStat = -1
    End Sub

    '==================== EVENTS ====================
    Public Event MasterRetrieved(ByVal Index As Integer, ByVal Value As Object)

    '==================== READ-ONLY PROPERTIES ====================
    Public ReadOnly Property AppDriver() As ggcAppDriver.GRider
        Get
            Return p_oApp
        End Get
    End Property

    Public ReadOnly Property EditMode() As xeEditMode
        Get
            Return p_nEditMode
        End Get
    End Property

    Public ReadOnly Property BranchCode() As String
        Get
            Return p_sBranchCd
        End Get
    End Property

    Public ReadOnly Property BranchName() As String
        Get
            Return p_sBranchNm
        End Get
    End Property

    '==================== MASTER ACCESSORS ====================
    Public Property Master(ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Return p_oDTMstr(0).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get
        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then

                Console.WriteLine(" index : " & strParm(Index))
                If (Index = 12) Then
                    p_oDTMstr(0).Item(Index) = DBNull.Value
                Else
                    p_oDTMstr(0).Item(Index) = value
                End If
            End If

        End Set
    End Property

    'Property Master(String)
    Public Property Master(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Return p_oDTMstr(0).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Console.WriteLine(" index : " & strParm(Index))
                p_oDTMstr(0).Item(Index) = value
            End If
        End Set
    End Property

    '==================== PARENT MODULE ====================
    Public Property Parent() As String
        Get
            Return p_sParent
        End Get
        Set(ByVal value As String)
            p_sParent = value
        End Set
    End Property

    '==================== DATA RETRIEVAL ====================
    ' Retrieves contracts based on application and contract status
    Public Function GetContractsByStatus(ByRef psMessage As String,
                                         branch As String,
                                         clientName As String,
                                         GOCASNo As String) As DataTable
        Dim loDT As New DataTable()



        Dim lsSQL As String =
            "SELECT DISTINCT " &
            "  b.sTransNox        AS sGOCASNox, " &
            "  b.sGOCASNoF, " &
            "  b.sClientNm, " &
            "  b.dTargetDt, " &
            "  b.nCrdtScrx, " &
            "  b.dTransact       AS dapplied, " &
            "  a.sTransNox       AS sContractNo, " &
            "  a.sBranchCd, " &
            "  a.dTransact       AS dContractDt, " &
            "  a.sClientID, " &
            "  a.sReferNox, " &
            "  a.sAcctNmbr, " &
            "  a.sSerialID, " &
            "  a.nAcctTerm, " &
            "  a.nDownPaym, " &
            "  a.nMonAmort, " &
            "  a.nRebatesx, " &
            "  a.nPenaltyx, " &
            "  a.dFirstPay, " &
            "  a.sRemarksx, " &
            "  a.cTranStat       AS cContractStat " &
            "FROM Credit_Online_Application b " &
            "LEFT JOIN MC_AR_Contract_Info a " &
            "  ON a.sReferNox = b.sTransNox " &
            "WHERE b.cTranStat = '2' " &
            "  AND a.cTranStat = '0' "

        If Not String.IsNullOrWhiteSpace(branch) Then
            lsSQL = AddCondition(lsSQL, "a.sBranchCd = " & strParm(branch))
        End If

        If Not String.IsNullOrWhiteSpace(clientName) Then
            lsSQL = AddCondition(lsSQL, "b.sClientNm LIKE " & strParm("%" & clientName & "%"))
        End If

        If Not String.IsNullOrWhiteSpace(GOCASNo) Then
            lsSQL = AddCondition(lsSQL, "a.sReferNox = " & strParm(GOCASNo))
        End If

        Console.WriteLine("Executing SQL: " & lsSQL)

        loDT = p_oApp.ExecuteQuery(lsSQL)

        ' Convert branch code to branch name
        If loDT IsNot Nothing AndAlso loDT.Rows.Count > 0 Then
            For Each row As DataRow In loDT.Rows
                If Not IsDBNull(row("sBranchCd")) Then
                    row("sBranchCd") = GetBranchName(row("sBranchCd").ToString())
                End If
            Next
        End If

        ' Set message and edit mode
        If loDT Is Nothing OrElse loDT.Rows.Count = 0 Then
            psMessage = "No record found."
        Else
            psMessage = ""
        End If

        Return loDT
    End Function

    '==================== LOOKUPS ====================
    ' Returns branch name based on branch code
    Public Function GetBranchName(branchCode As String) As String
        Dim BranchName As String = branchCode

        Dim lsSQL As String =
            "SELECT sBranchNm FROM Branch WHERE sBranchCd = " & strParm(branchCode)

        Dim dt As DataTable = p_oApp.ExecuteQuery(lsSQL)

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            BranchName = dt.Rows(0)("sBranchNm").ToString()
        End If

        Return BranchName
    End Function

    ' Returns client company name
    Public Function GetClientName(clientID As String) As String
        Dim loClient As New ggcClient.Client(p_oApp)
        loClient.Parent = "MCARContract"

        If Not loClient.OpenClient(clientID) Then
            Return "Client not found."
        End If

        Return loClient.Master("sCompnyNm")
    End Function

    ' Returns PN value from LR Master
    Public Function GetPNValue(accountNumber As String) As String
        Dim loMaster As New LRMaster(p_oApp)
        loMaster.Parent = "MCARContract"

        If Not loMaster.OpenTransaction(accountNumber) Then
            MessageBox.Show("No Record Found", "Information",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return "0.00"
        End If

        Return loMaster.Master("")
    End Function

    ' Returns engine and frame number
    Public Function GetSerialInfo(serialID As String, branchcd As String) As Tuple(Of String, String)
        Dim engineNo As String = ""
        Dim frameNo As String = ""
        Dim msg As String = ""

        Dim lsSQL As String =
            " SELECT sEngineNo, sFrameNox,cLocation,cSoldStat " &
            " FROM MC_Serial " &
            " WHERE sSerialID = " & strParm(serialID) &
            " AND sBranchCd =  " & strParm(branchcd)


        Dim dt As DataTable = p_oApp.ExecuteQuery(lsSQL)


        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            If (dt.Rows(0)("cLocation") = "1" And dt.Rows(0)("cSoldStat") = "0") Then
                engineNo = dt.Rows(0)("sEngineNo").ToString()
                frameNo = dt.Rows(0)("sFrameNox").ToString()
            Else
                engineNo = dt.Rows(0)("sEngineNo").ToString()
                frameNo = dt.Rows(0)("sFrameNox").ToString()
                msg = "We couldn’t find this serial in the branch inventory. Please double-check the serial number."
                ' Show message box
                MessageBox.Show(msg, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            MessageBox.Show("No Record Found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If


        Return Tuple.Create(engineNo, frameNo)
    End Function

    ' Returns PN value from MC AR Master
    Public Function GetMCARMaster(accountNo As String) As String
        Dim pnValue As String = ""

        Dim lsSQL As String =
            "SELECT nPNValuex " &
            "FROM MC_AR_Master " &
            "WHERE sAcctNmbr = " & strParm(accountNo)

        Dim dt As DataTable = p_oApp.ExecuteQuery(lsSQL)

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            pnValue = dt.Rows(0)("nPNValuex").ToString()
        End If

        Return pnValue
    End Function

    '==================== SQL BUILDER ====================
    Private Function getSQ_Master() As String
        Return "SELECT" &
               "  sTransNox" &
               ", sBranchCd" &
               ", dTransact" &
               ", sClientID" &
               ", sReferNox" &
               ", sAcctNmbr" &
               ", sSerialID" &
               ", nDownPaym" &
               ", nAcctTerm" &
               ", nMonAmort" &
               ", nRebatesx" &
               ", nPenaltyx" &
               ", dFirstPay" &
               ", sRemarksx" &
               ", cTranStat" &
               ", sModified" &
               ", dModified" &
               " FROM " & p_sMasTable & " a"
    End Function

    '==================== TRANSACTION METHODS ====================
    Public Function OpenTransaction(ByVal fsTransNox As String) As Boolean
        Dim lsSQL As String =
            AddCondition(getSQ_Master, "a.sTransNox = " & strParm(fsTransNox))

        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function

    Public Function UpdateTransaction(ByVal fsTransNox As String) As Boolean
        If p_nEditMode <> xeEditMode.MODE_READY Then
            MsgBox("Invalid Edit Mode detected!",
                   MsgBoxStyle.OkOnly + MsgBoxStyle.Critical,
                   p_sMsgHeadr)
            Return False
        End If

        If String.IsNullOrEmpty(fsTransNox) Then
            MsgBox("Transaction number is empty!",
                   MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation,
                   p_sMsgHeadr)
            Return False
        End If

        p_nEditMode = xeEditMode.MODE_UPDATE
        Return True
    End Function

    Public Function SaveTransaction() As Boolean
        If p_nEditMode <> xeEditMode.MODE_UPDATE Then
            MsgBox("Invalid Edit Mode detected!",
                   MsgBoxStyle.OkOnly + MsgBoxStyle.Critical,
                   p_sMsgHeadr)
            Return False
        End If

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            Dim lsSQL As String =
                "UPDATE " & p_sMasTable &
                " SET   nMonAmort = " & strParm(p_oDTMstr(0).Item("nMonAmort")) &
                ", nRebatesx = " & strParm(p_oDTMstr(0).Item("nRebatesx")) &
                ", nPenaltyx = " & strParm(p_oDTMstr(0).Item("nPenaltyx")) &
                ", cTranStat = " & strParm(CStr(xeTranStat.TRANS_OPEN)) &
                ", sModified = " & strParm(p_oApp.UserID) &
                ", dModified = " & dateParm(p_oApp.SysDate) &
                " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))

            Call p_oApp.Execute(lsSQL, p_sMasTable)
            If p_sParent = "" Then
                p_oApp.CommitTransaction()
                MessageBox.Show("Transaction Save Succesfully!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                p_nEditMode = xeEditMode.MODE_READY
            End If
            Return True

        Catch ex As Exception
            If p_sParent = "" Then p_oApp.RollBackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function ApproveTransaction(transaction As String, psGocasTransNo As String) As Boolean
        If Not p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            MsgBox("Invalid Edit Mode detected!",
                   MsgBoxStyle.OkOnly + MsgBoxStyle.Critical,
                   p_sMsgHeadr)
            Return False
        End If
        If String.IsNullOrEmpty(transaction) Then
            MsgBox("Please load a transaction before proceeding.", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical,
                   p_sMsgHeadr)
            Return False  ' stop execution
        End If

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            Dim lsSQL As String =
                "UPDATE " & p_sMasTable &
                " SET   cTranStat = " & strParm(CStr(xeTranStat.TRANS_CLOSED)) &
                ", sModified = " & strParm(p_oApp.UserID) &
                ", dModified = " & dateParm(p_oApp.SysDate) &
                " WHERE sTransNox = " & strParm(transaction)

            Call p_oApp.Execute(lsSQL, p_sMasTable)
            If p_sParent = "" Then
                p_oApp.CommitTransaction()
                MessageBox.Show("Transaction Posted Succesfully!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                If (RMJExecuteL("D:\GGC_Java_Systems", "gocas-export.bat", psGocasTransNo) <> 0) Then
                    MsgBox("Unable to export GOCAS Form.", vbCritical, "Warning")
                Else
                    MsgBox("GOCAS Form export successful.", vbInformation, "Information")
                End If

                p_nEditMode = xeEditMode.MODE_READY
            End If
            Return True

        Catch ex As Exception
            If p_sParent = "" Then p_oApp.RollBackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function ExportTransaction(transaction As String, psGocasTransNo As String, transStat As String) As Boolean
        If Not p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            MsgBox("Invalid Edit Mode detected!",
                   MsgBoxStyle.OkOnly + MsgBoxStyle.Critical,
                   p_sMsgHeadr)
            Return False
        End If
        If String.IsNullOrEmpty(transaction) Then
            MsgBox("Please load a transaction before proceeding.", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical,
                   p_sMsgHeadr)
            Return False  ' stop execution
        End If

        If Not transStat = xeTranStat.TRANS_CLOSED Then
            MsgBox("Export is allowed only for approved transactions.", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical,
                   p_sMsgHeadr)
            Return False  ' stop execution
        End If

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            If (RMJExecuteL("D:\GGC_Java_Systems", "gocas-export.bat", psGocasTransNo) <> 0) Then
                MsgBox("Unable to export GOCAS Form.", vbCritical, "Warning")
            Else
                MsgBox("GOCAS Form export successful.", vbInformation, "Information")
                p_nEditMode = xeEditMode.MODE_UNKNOWN
            End If

            Return True

        Catch ex As Exception
            If p_sParent = "" Then p_oApp.RollBackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Public Sub resetEditmode()
        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Public Function SearchTransaction(
                        ByVal fsValue As String,
                        branch As String,
                        clientName As String,
                        GOCASNo As String,
                        isEntry As Boolean,
                        Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String

        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMstr(0).Item("sTransNox") Then Return True
            End If
        End If
        'Initialize SQL filter
        If isEntry Then
            lsSQL = AddCondition(getSQ_Browse, "a.cTranStat  = '0'")
        Else
            lsSQL = AddCondition(getSQ_Browse, "a.cTranStat  = '1'")
        End If
        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sReferNox LIKE " & strParm(fsValue)
        Else
            lsFilter = "b.sCompnyNm LIKE " & strParm(fsValue & "%")
        End If

        If Not String.IsNullOrWhiteSpace(branch) Then
            lsSQL = AddCondition(lsSQL, "a.sBranchCd = " & strParm(branch))
        End If

        If Not String.IsNullOrWhiteSpace(clientName) Then
            lsSQL = AddCondition(lsSQL, "b.sClientNm LIKE " & strParm("%" & clientName & "%"))
        End If

        If Not String.IsNullOrWhiteSpace(GOCASNo) Then
            lsSQL = AddCondition(lsSQL, "a.sReferNox = " & strParm(GOCASNo))
        End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , fsValue _
                                        , "sTransNox»xClientNm»dTransact»sReferNox»xBranchNm" _
                                        , "Transaction No»Client»Date»GOCAS No»Branch" _
                                        , "" _
                                        , "a.sTransNox»b.sCompnyNm»a.dTransact»a.sReferNox»c.sBranchNm" _
                                        , IIf(fbByCode, 0, 1))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenTransaction(loDta.Item("sTransNox"))
            p_nEditMode = xeEditMode.MODE_READY
        End If
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT" &
                    " a.sTransNox" &
                    ", a.sBranchCd" &
                    ", a.dTransact" &
                    ", a.sClientID" &
                    ", a.sReferNox" &
                    ", a.sAcctNmbr" &
                    ", a.sSerialID" &
                    ", a.nDownPaym" &
                    ", a.nAcctTerm" &
                    ", a.nMonAmort" &
                    ", a.nRebatesx" &
                    ", a.nPenaltyx" &
                    ", a.dFirstPay" &
                    ", a.sRemarksx" &
                    ", a.cTranStat" &
                    ", a.sModified" &
                    ", a.dModified" &
                    ", b.sCompnyNm xClientNm" &
                    ", c.sBranchNm xBranchNm" &
                    " FROM " & p_sMasTable & " a" &
                    " LEFT JOIN Client_Master b ON a.sClientID = b.sClientID" &
                    " LEFT JOIN Branch c ON a.sBranchCd = c.sBranchCd" &
              " WHERE a.sClientID = b.sClientID"
    End Function


    Public Sub SearchBranch(ByVal fsValue As String _
                          , ByVal fbIsCode As Boolean _
                          , ByVal fbIsSrch As Boolean)


        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_sBranchCd And fsValue <> "" Then Exit Sub
        Else
            If fsValue = p_sBranchNm And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" &
                       "  a.sBranchCD" &
                       ", a.sBranchNm" &
               " FROM Branch a" &
               IIf(fbIsCode = False, " WHERE a.cRecdStat = '1'", "")

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sBranchCD»sBranchNm" _
                                             , "ID»Company",
                                             , "a.sBranchCD»a.sBranchNm" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_sBranchCd = ""
                p_sBranchNm = ""
            Else
                p_sBranchCd = loRow.Item("sBranchCD")
                p_sBranchNm = loRow.Item("sBranchNm")
            End If
            Exit Sub
        End If

        If fsValue <> "" Then
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sBranchCD = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sBranchNm = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_sBranchCd = ""
            p_sBranchNm = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_sBranchCd = loDta(0).Item("sBranchCD")
            p_sBranchNm = loDta(0).Item("sBranchNm")
        End If
    End Sub


End Class
