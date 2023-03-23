<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Statistiken erstellen.aspx.cs" Inherits="Excel_Hochladen_und_Einlesen.Statistiken_erstellen" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        Konsument :
        <asp:TextBox ID="txt_Konsument" runat="server"></asp:TextBox>
        &nbsp;(z.B. AT0030000000000000000000000CON001)<br />
        <br />
        <br />
        <asp:Calendar ID="cal_Von" runat="server"></asp:Calendar>
        <br />
        <asp:Calendar ID="cal_Bis" runat="server"></asp:Calendar>
        <br />
        <br />
        <br />
        <div>
            <asp:Button ID="btn_Statistik" runat="server" OnClick="btn_Statistik_Click" Text="Statistik erstellen" />
        </div>
    </form>
</body>
</html>
