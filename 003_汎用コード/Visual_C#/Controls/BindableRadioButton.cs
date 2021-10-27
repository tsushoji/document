using System;
using System.Windows.Forms;

namespace BitOperationWinApp.Controls
{
    public class BindableRadioButton : RadioButton
    {
        public BindableRadioButton()
        {
            AutoCheck = false;
        }
        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);

            this.Checked = !this.Checked;
        }
    }
}
