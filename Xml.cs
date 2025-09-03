using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToolkitAddIn
{
    public class Xml
    {
        public string footer = "</tabs></ribbon></customUI>";

        public string header =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"Ribbon_Load\"><ribbon><tabs>";

        public Tab[] tabs;

        public Xml(params Tab[] tabs)
        {
            this.tabs = tabs;
        }
        public string ToXml()
        {
            StringBuilder builder = new StringBuilder(2048 * tabs.Length);
            builder.Append(header);
            foreach (Tab tab in tabs) builder.Append(tab.ToXml());
            builder.Append(footer);
            return builder.ToString();
        }

        public struct Tab
        {
            public string id;
            public string label;
            public Group[] groups;

            /// <summary>
            /// 生成Ribbon XML Tab标签
            /// </summary>
            /// <param name="id">对应 id</param>
            /// <param name="label">对应 label</param>
            /// <param name="groups">对应 group成员，可同时传入多个Group</param>
            public Tab(string id, string label, params Group[] groups)
            {
                this.id = id;
                this.label = label;
                this.groups = groups;
            }

            /// <summary>
            /// 生成Ribbon XML Tab字符串
            /// </summary>
            /// <returns></returns>
            public string ToXml()
            {
                StringBuilder builder = new StringBuilder(512 * groups.Length);
                builder.Append($@"<tab id=""{id}"" label=""{label}"">");
                foreach (Group group in groups) builder.Append(group.ToXml());
                builder.Append("</tab>");
                return builder.ToString();
            }
            
            public Tab AddGroup(params Group[] groups)
            {
                Group[] newGroups = new Group[this.groups.Length + groups.Length];
                Array.Copy(this.groups,0,newGroups,0,this.groups.Length);
                Array.Copy(groups, 0, newGroups, this.groups.Length, groups.Length);
                return this;
            }
        }

        public struct Group
        {
            public string id;
            public string label;
            public IControl[] controls;

            /// <summary>
            /// 创建Group
            /// </summary>
            /// <param name="id">对应 id</param>
            /// <param name="label">对应 label</param>
            /// <param name="controls">对应 control成员，可同时传入多个Control
            /// </param>
            public Group(string id, string label, params IControl[] controls)
            {
                this.id = id;
                this.label = label;
                this.controls = controls;
            }

            public Group(string id, string label, params IControl[][] controlArrArr)
            {
                this.id = id;
                this.label = label;
                this.controls = new IControl[controlArrArr.Sum(controlArr => controlArr?.Length ?? 0)];
                int offset = 0;
                foreach (IControl[] controlArr in controlArrArr)
                {
                    if (controlArr == null) continue;
                    Array.Copy(controlArr, 0, this.controls, offset, controlArr.Length);
                    offset += controlArr.Length;
                }
            }

            public Group AddControl(params IControl[] controls)
            {
                var newControls = new IControl[this.controls.Length + controls.Length];
                Array.Copy(this.controls, newControls, this.controls.Length);
                Array.Copy(controls, 0, newControls, this.controls.Length, controls.Length);
                this.controls = newControls;
                return this;
            }

            /// <summary>
            /// 生成Ribbon XML Group字符串
            /// </summary>
            public string ToXml()
            {
                StringBuilder builder = new StringBuilder(512);
                builder.Append($@"<group id=""{id}"" label=""{label}"">");
                foreach (IControl control in controls) builder.Append(control.ToXml());
                builder.Append("</group>");
                return builder.ToString();
            }
        }

        public interface IControl
        {
            string ToXml();
        }

        public struct Button : IControl
        {
            public string id;
            public string label;
            public string onAction;
            public string imageMso;
            public string size;

            public Button(string id, string label, string onAction, string size = null, string imageMso = null)
            {
                this.id = id;
                this.label = label;
                this.onAction = onAction;
                this.size = size;
                this.imageMso = imageMso;
            }

            public string ToXml()
            {
                string sizeXml = size == null ? "" : $"size=\"{size}\"";
                string imageXml = imageMso == null ? "" : $"imageMso=\"{imageMso}\"";
                return $"<button id=\"{id}\" label=\"{label}\" onAction=\"{onAction}\" {sizeXml} {imageXml}/>";
            }
        }

        public struct CheckBox : IControl
        {
            public string id;
            public string label;
            public string onAction;
            public string getPressed;
            public CheckBox(string id, string label, string onAction, Dictionary<string, bool> isChecks, bool isCheck, string getPressed = "GetChecked")
            {
                this.id = id;
                this.label = label;
                this.onAction = onAction;
                this.getPressed = getPressed;
                isChecks[this.id] = isCheck;
            }

            public string ToXml()
            {
                return $"<checkBox id=\"{id}\" label=\"{label}\" onAction=\"{onAction}\" getPressed=\"{getPressed}\" />";
            }
        }

        public struct SplitButton : IControl
        {
            public string id;
            public string label;
            public string onAction;
            public string imageMso;
            public IControl[] controls;

            public SplitButton(string id, string label, string onAction, string imageMso, params IControl[] controls)
            {
                this.id = id;
                this.label = label;
                this.onAction = onAction;
                this.imageMso = imageMso;
                this.controls = controls;
            }

            public string ToXml()
            {
                StringBuilder builder = new StringBuilder(1024);
                builder.Append($@"<splitButton id=""SplitCheck{id}"" size=""large"" ><button id=""{id}"" label=""{label}"" onAction=""{onAction}"" imageMso=""{imageMso}""/><menu>");
                foreach (IControl control in controls) builder.Append(control.ToXml());
                builder.Append("</menu></splitButton>");
                return builder.ToString();
            }

            public SplitButton AddControl(params IControl[] controls)
            {
                IControl[] newControls = new IControl[this.controls.Length + controls.Length];
                Array.Copy(this.controls, 0, newControls, 0, this.controls.Length);
                Array.Copy(controls, 0, newControls, this.controls.Length, controls.Length);
                this.controls = newControls;
                return this;
            }
        }

        public struct SplitCheckBox : IControl
        {
            public string id;
            public string label;
            public string onAction;
            public string imageMso;
            public CheckBox[] checks;
            public string[] checksId;

            public SplitCheckBox(string id, string label, string onAction, string imageMso, params CheckBox[] checks)
            {
                this.id = id;
                this.label = label;
                this.onAction = onAction;
                this.imageMso = imageMso;
                this.checks = checks;
                this.checksId = (string[])checks.Select(x => x.id);
            }

            public string ToXml()
            {
                StringBuilder builder = new StringBuilder(512);
                builder.Append($@"<splitButton id=""SplitCheck{id}"" label="""" size=""large""  imageMso=""{imageMso}"">");
                builder.Append($@"<button id=""{id}"" label=""{label}"" onAction=""{onAction}""/><menu>");
                foreach (CheckBox check in checks) builder.Append(check.ToXml());
                builder.Append("</menu></splitButton>");
                return builder.ToString();
            }
        }

        public struct Separator : IControl
        {
            public string id;

            public Separator(string id)
            {
                this.id = id;
            }

            public string ToXml()
            {
                return $"<separator id=\"{id}\" />";
            }
        }
    }
}