using System;

namespace BtFiller2.Services
{
    /// <summary>
    /// 构建有效期守卫：BuildUtc + Ttl。
    /// BuildUtc 取自编译期由 MSBuild 目标 GenerateBuildInfo 写入的
    /// <see cref="BuildInfo.BuildTicksUtc"/>，因此每次重新出包会自动刷新，
    /// 无需手工维护常量。若改为手工常量，请在此处更新 BuildUtc。
    /// 到期后 <see cref="IsExpired"/> 返回 true，样机将无法初始化。
    /// 整个过程（之前 / 之中 / 之后）不产生任何提示。
    /// </summary>
    internal static class LicenseGuard
    {
        /// <summary>有效期（分钟）。518400 分钟 = 360 天。</summary>
        internal const double TtlMinutes = 518400d;

        /// <summary>构建时刻（UTC）。由编译期生成，随每次出包自动更新。</summary>
        internal static DateTime BuildUtc => new DateTime(BuildInfo.BuildTicksUtc, DateTimeKind.Utc);

        /// <summary>是否已超出有效期（含时钟被回拨的异常情况）。</summary>
        internal static bool IsExpired()
        {
            // 实现细节未公开（算法核心）
            throw new NotImplementedException("该方法的实现不在开源范围内");
        }

        /// <summary>到期则静默退出进程，不弹出任何提示。</summary>
        internal static void EnforceOrExit()
        {
            // 实现细节未公开（算法核心）
            throw new NotImplementedException("该方法的实现不在开源范围内");
        }
    }
}
