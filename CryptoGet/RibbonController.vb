Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI

Namespace Ribbon
    <ComVisible(True)>
    Public Class RibbonController
        Inherits ExcelRibbon

        Public Overrides Function GetCustomUI(ByVal RibbonID As String) As String
            Return "
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab idMso='TabHome'>
            <group id='cryptoGeT' label='...'>
              <button id='button1' label='Crypto Get'  size='large' imageMso='HappyFace' onAction='OnButtonPressed'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>"
        End Function

        Public Sub OnButtonPressed(ByVal control As IRibbonControl)
            Dim FRM As New Form1
            FRM.ShowDialog()
        End Sub
    End Class
End Namespace
