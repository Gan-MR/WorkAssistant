using System;
using System.Deployment.Application;
using System.Windows.Forms;

namespace 工作助手.Function
{
    internal class UpDate
    {
        public static void InstallUpdateSyncWithInfo()
        {
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;

                UpdateCheckInfo info;
                try
                {
                    info = ad.CheckForDetailedUpdate();

                }
                catch (DeploymentDownloadException dde)
                {
                    MessageBox.Show("无法下载 \n\n请检查网络通讯: " + dde.Message);
                    return;
                }
                catch (InvalidDeploymentException ide)
                {
                    MessageBox.Show("无法检测新版本，应用程序损坏: " + ide.Message);
                    return;
                }
                catch (InvalidOperationException ioe)
                {
                    MessageBox.Show("无法更新，程序类型错误: " + ioe.Message);
                    return;
                }

                if (info.UpdateAvailable)
                {
                    Boolean doUpdate = true;

                    if (!info.IsUpdateRequired)
                    {
                        DialogResult dr = MessageBox.Show("有更新。您想现在更新应用程序吗？", "有可用更新", MessageBoxButtons.OKCancel);
                        if (!(DialogResult.OK == dr))
                        {
                            doUpdate = false;
                        }
                    }
                    else
                    {
                        // Display a message that the app MUST reboot. Display the minimum required version.
                        MessageBox.Show("已检测到强制更新 " +
                            "version to version " + info.MinimumRequiredVersion.ToString() +
                            " 应用程序现在将安装更新并重新启动。 ",
                            "有可用更新", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    if (doUpdate)
                    {
                        try
                        {
                            ad.Update();
                            MessageBox.Show("更新完成，将重新启动应用程序。");
                            Application.Restart();
                        }
                        catch (DeploymentDownloadException dde)
                        {
                            MessageBox.Show("无法安装最新版本的应用程序。请检查您的网络连接，或者稍后再试。错误: " + dde);
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("没有找到更新", "寄！", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("应用程序当前无法更新。", "寄！", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }
    }
}
