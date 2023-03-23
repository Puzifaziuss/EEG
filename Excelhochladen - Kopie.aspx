<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Excelhochladen.aspx.cs" Inherits="Excel_Hochladen_und_Einlesen.Excelhochladen" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    
    <style type="text/css">
        .auto-style1 {
            height: 12px;
            width: 64px;
        }
        .auto-style2 {
            height: 23px;
        }
        .auto-style3 {
            width: 442px;
        }
        .auto-style4 {
            height: 12px;
            width: 442px;
        }
        .auto-style5 {
            height: 23px;
            width: 442px;
        }
        .auto-style6 {
            width: 64px;
        }
        .auto-style7 {
            height: 23px;
            width: 64px;
        }
        .auto-style8 {
            width: 157px;
        }
        .auto-style9 {
            height: 23px;
            width: 157px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br />
            &nbsp;<asp:Button ID="btnLastFormat" runat="server" Text="Letztes Format verwenden" OnClick="btnLastFormat_Click" />
            &nbsp;<table style="width:100%;">
                <tr>
                    <td class="auto-style3">&nbsp;</td>
                    <td class="auto-style6"> &nbsp;</td>
                    <td class="auto-style8">&nbsp;</td>
                     <td> &nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td class="auto-style3">Anzahl Konsumenten:</td>
                    <td class="auto-style6"> <asp:TextBox ID="txtAnzCon" runat="server" Height="21px" Width="47px"></asp:TextBox>
                    </td>
                    <td class="auto-style8">Spalten pro Konsument: </td>
                     <td> <asp:TextBox ID="txtConCol" runat="server" Width="49px"></asp:TextBox>
                    </td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td class="auto-style3">Anzahl Erzeuger:</td>
                    <td class="auto-style6"> <asp:TextBox ID="txtAnzGen" runat="server" Width="46px"></asp:TextBox>
                    </td>
                    <td class="auto-style8">Spalten pro Erzeuger: </td>
                     <td> <asp:TextBox ID="txtGenCol" runat="server" Width="52px"></asp:TextBox>
                    </td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td class="auto-style4">Welche Tabelle der -xlsx-Datei soll eingelesen werden? (1, 2 o. 3 ...) </td>
                    <td class="auto-style1">
            <asp:TextBox ID="txtNumTab" runat="server" Width="44px"></asp:TextBox>
                    </td>
                  
                </tr>
                 <tr>
                    <td class="auto-style5">
                        <asp:CustomValidator ID="cv_CheckFields" runat="server" EnableClientScript="False" ErrorMessage="Bitte geben Sie das ganz Format der Excel-Tabelle ein" OnServerValidate="cv_CheckFields_ServerValidate"></asp:CustomValidator>
                     </td>
                    <td class="auto-style7"></td>
                    <td class="auto-style9"></td>
                      <td class="auto-style2"></td>
                    <td class="auto-style2"></td>
                    <td class="auto-style2"></td>
                </tr>
                 <tr>
                    <td class="auto-style5"></td>
                    <td class="auto-style7"></td>
                    <td class="auto-style9"></td>
                      <td class="auto-style2"></td>
                    <td class="auto-style2"></td>
                    <td class="auto-style2"></td>
                </tr>
                 <tr>
                    <td class="auto-style5">&nbsp;</td>
                    <td class="auto-style7">&nbsp;</td>
                    <td class="auto-style9">&nbsp;</td>
                      <td class="auto-style2">&nbsp;</td>
                    <td class="auto-style2">&nbsp;</td>
                    <td class="auto-style2">&nbsp;</td>
                </tr>
                 <tr>
                    <td class="auto-style5"></td>
                    <td class="auto-style7"></td>
                    <td class="auto-style9"></td>
                      <td class="auto-style2"></td>
                    <td class="auto-style2"></td>
                    <td class="auto-style2"></td>
                </tr>
            </table>
            <br />
            <br />
            Bitte laden Sie jeweils eine Datei pro Monat hoch.<br />
            <br />
            <asp:FileUpload ID="fuReport" runat="server" />
            &nbsp;&nbsp;&nbsp;&nbsp;
            <asp:RequiredFieldValidator ID="rfvUpload" runat="server" ControlToValidate="fuReport" EnableClientScript="False" ErrorMessage="Bitte wählen sie eine .xlsx-Datei aus"></asp:RequiredFieldValidator>
            <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <br />
            <br />
            <asp:Label ID="lblInfo" runat="server"></asp:Label>
            <br />
            <br />
            <br />
            <br />
            <asp:Button ID="btnEinlesen" runat="server" OnClick="btnEinlesen_Click" Text="Einlesen" />
            <br />
            <br />
            <br />
        </div>
    </form>
    <p>
        s</p>
</body>
</html>
