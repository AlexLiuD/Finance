﻿using System.Web;
using System.Web.Optimization;

namespace Finance
{
    public class BundleConfig
    {
        // 有关绑定的详细信息，请访问 http://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                        "~/Scripts/jquery-{version}.js"));

            bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
                        "~/Scripts/jquery.validate*"));

            // 使用要用于开发和学习的 Modernizr 的开发版本。然后，当你做好
            // 生产准备时，请使用 http://modernizr.com 上的生成工具来仅选择所需的测试。
            bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
                        "~/Scripts/modernizr-*"));

            bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
                        "~/Scripts/bootstrap.js",
                        "~/Scripts/respond.js",
                        "~/Scripts/bootstrap-datetimepicker.js",
                        "~/Scripts/bootstrap-datetimepicker.zh-CN.js"));
            //返回顶部
            bundles.Add(new ScriptBundle("~/bundles/top").Include(
                        "~/Scripts/top.js"));
            //bootstrap-table
            bundles.Add(new ScriptBundle("~/bundles/bootstrap-table").Include(
                        "~/Scripts/bootstrap-table/bootstrap-table.js",
                        "~/Scripts/bootstrap-table/locale/bootstrap-table-zh-CN.js",
                        //"~/Scripts/bootstrap-table/extensions/export/tableExport.js",
                        //"~/Scripts/bootstrap-table/extensions/export/jquery.base64.js",
                        //"~/Scripts/bootstrap-table/extensions/export/bootstrap-table-export.js",
                        //"~/Scripts/bootstrap-table/extensions/filter/table-filter.js",
                        //"~/Scripts/bootstrap-table/extensions/filter/bs-table.js",
                        //"~/Scripts/bootstrap-table/extensions/filter/bootstrap-table-filter.js",
                        "~/Scripts/bootstrap-select.js"));

            bundles.Add(new StyleBundle("~/Content/css").Include(
                       "~/Content/bootstrap.css",
                       "~/Content/site.css",
                       "~/Content/bootstrap-datetimepicker.css",
                       "~/Content/top.css",
                       "~/Content/bootstrap-table.css",
                       "~/Content/bootstrap-select.css",
                       "~/Content/bootstrap-table-filter.css"));
        }
    }
}
