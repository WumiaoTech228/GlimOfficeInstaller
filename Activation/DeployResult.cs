using System.Collections.Generic;

namespace GOI.Activation
{
    /// <summary>
    /// 部署/卸载操作的返回结果。
    /// </summary>
    public class DeployResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string Phase { get; set; }
        public List<string> Steps { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
    }
}
