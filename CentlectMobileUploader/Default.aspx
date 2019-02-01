<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="CentlectMobileUploader._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1>Centlec Mobile</h1>
    </div>

    <div class="row">
       <h2>Please upload a Excel Spreadsheet</h2>
    </div>
    <hr>
    <div class="row">
        <div class="col-md-3">
            <asp:FileUpload ID="FileUpload1" runat="server" />
        </div>        
        <div class="col-md-9">
            <asp:Button ID="Button1" runat="server" class="btn btn-success btn-sm" Text="Upload file" />
        </div>
    </div>
    <asp:Label ID="lblResult" Visible="false" runat="server" Text=""></asp:Label>
</asp:Content>
