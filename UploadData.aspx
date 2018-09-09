<%@ Page Title="" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeFile="UploadData.aspx.cs" Inherits="UploadData" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="Server">
    <script src="Scripts/jquery-1.11.3.min.js" type="text/javascript"></script>
    <script src="Scripts/select2.min.js" type="text/javascript"></script>
    <link href="Styles/select2.min.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript">
        $(function () {
            $(".clsExcelColum").select2({
                placeholder: "--Select--",
                allowClear: true
            });
        });
    </script>

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="Server">
    <h2>
        Upload Data
    </h2>
    <hr />
    <div id="divselectFolder">
        Folder Path* :
        <asp:TextBox ID="txtfolderPath" runat="server"></asp:TextBox>
        <asp:RequiredFieldValidator ErrorMessage="Please provide folder path" ControlToValidate="txtfolderPath"
            runat="server" Display="Dynamic" ValidationGroup="valSelectfile" ForeColor="Red" />
        <asp:Button Text="Get Files" ID="btnGetFiles" runat="server" CausesValidation="true"
            ValidationGroup="valSelectfile" OnClick="btnGetFiles_Click" />
    </div>
    <div id="divFileList">
        Data Read: <span style="background-color: Green;" class="ColorSpan">&nbsp;&nbsp;</span>
        Excel Rejected: <span style="background-color: Red;" class="ColorSpan">&nbsp;&nbsp;</span>
        Process Pending: <span style="background-color: Gray" class="ColorSpan">&nbsp;&nbsp;</span>
        In Process: <span style="background-color: Maroon" class="ColorSpan">&nbsp;&nbsp;</span>
        <br />
        <asp:GridView runat="server" ID="grdFileQuea" AutoGenerateColumns="false" OnRowDataBound="grdFileQuea_RowDataBound">
            <Columns>
                <asp:TemplateField HeaderText="File ID">
                    <ItemTemplate>
                        <asp:Label Text='<%# Eval("FileId") %>' runat="server" />
                        <%--<asp:HiddenField ID="hdnFileStauts" Value='<%# Eval("FileStatus") %>' runat="server" />--%>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Excel File">
                    <ItemTemplate>
                        <asp:Label ID="Label1" Text='<%# Eval("FilePath") %>' runat="server" Font-Bold="true" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Excel Sheet">
                    <ItemTemplate>
                        <asp:Label ID="Label1" Text='<%# Eval("SheetName") %>' runat="server" Font-Bold="true" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
    </div>
    <br />
    <div id="divcolumMappingUpload">
        <asp:Label Text="" ID="lblWarning" ForeColor="Red" Font-Bold="true" runat="server" />
        <br />
        <asp:GridView runat="server" ID="grdColumnMapping" AutoGenerateColumns="false" OnRowDataBound="grdColumnMapping_RowDataBound"
            DataKeyNames="ProcessFileId,FilePath,SheetName">
            <Columns>
                <asp:TemplateField HeaderText="Sr no.">
                    <ItemTemplate>
                        <%# Container.DataItemIndex + 1 %>
                        <asp:HiddenField ID="hdnProcessFieldId" Value='<%# Eval("ProcessFileId") %>' runat="server" />
                        <asp:HiddenField ID="hdnFilePath" Value='<%# Eval("FilePath") %>' runat="server" />
                        <asp:HiddenField ID="hdnSheetName" Value='<%# Eval("SheetName") %>' runat="server" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Column Name">
                    <ItemTemplate>
                        <asp:Label ID="lblMasterColum" Text='<%# Eval("masterColumnName") %>' runat="server" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Excel Column Name">
                    <ItemTemplate>
                        <asp:DropDownList runat="server" ID="drdwExcelColumn" CssClass="clsExcelColum" TabIndex="1">
                        </asp:DropDownList>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Default Data">
                    <ItemTemplate>
                        <asp:TextBox runat="server" ID="txtDefaultData" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Error Msg">
                    <ItemTemplate>
                        <asp:Label Text="" ID="lblErrMsg" runat="server" ForeColor="Red" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        <br />
        <div style="color: Red">
            <asp:Label ID="lblValidationError" runat="server" />
        </div>
        <br />
        <asp:RadioButtonList runat="server" ID="chkOverride">
            <asp:ListItem Text="Delete and Add" Value="True" />
            <asp:ListItem Text="Add" Value="False" Selected="True" />
        </asp:RadioButtonList>
        <asp:Button Text="Upload Data and move to next" runat="server" ID="btnUploadData"
            OnClick="btnUploadData_Click" />
        <asp:Button Text="Reject Excel and move to next" runat="server" ID="btnRejectExcel"
            OnClick="btnRejectExcel_Click" />
        <asp:HiddenField ID="hdnProcessFieldId" Value="" runat="server" />
        <asp:HiddenField ID="hdnFilePath" Value="" runat="server" />
        <asp:HiddenField ID="hdnSheetName" Value="" runat="server" />
    </div>
</asp:Content>
