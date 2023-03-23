<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Rechnung erstellen.aspx.cs" Inherits="Excel_Hochladen_und_Einlesen.Rechnung_erstellen" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            Rechnung erstellen<br />
            <br />
            Mitglied:
            <asp:DropDownList ID="DropDownList1" runat="server">
            </asp:DropDownList>
            <br />
            <br />
            <br />
            Zeitraum:<br />
            <br />
            Jahr:<br />
            <asp:DropDownList ID="DropDownList3" runat="server">
            </asp:DropDownList>
            <br />
            <br />
            Quartal: <br />
            <asp:DropDownList ID="DropDownList2" runat="server">
            </asp:DropDownList>
            <br />
            <br />
            <asp:Button ID="btn_RechnungErstellen" runat="server" OnClick="btn_RechnungErstellen_Click" Text="Rechnung Erstellen" />
            <br />
        </div>
    </form>
</body>
</html>
