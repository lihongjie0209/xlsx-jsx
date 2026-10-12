# xlsx-jsx Agent 工作约定

目标是用 Rust CLI 将 JSX/TSX 描述编译为 XLSX，并将支持子集的外部 XLSX
逆向为可重新编译的 JSX。Node.js 不是运行时或编译后端。

## 核心不变量

- 只开放正向 writer 与外部 XLSX reverse 均可表达的能力；不得依赖私有 manifest
  证明逆向可用。
- 逆向不得静默丢弃公式、样式、单元格类型或工作簿结构；超出无损范围时明确报错，
  指出原因和可执行修复建议。
- CLI 面向 Agent：命令和输出确定、可脚本化；`spec` 输出规范，`validate` 不写 XLSX。
- 保持增量构建与缓存；不要为了排错清空全局 Cargo 缓存或整个 `target`。

## 功能批次

1. 覆盖审计：以锁定版本 `rust_xlsxwriter` 的 writer API 和实际 OOXML 为准，更新
   `docs/xlsxwriter-coverage.md`，分类 covered、writer 缺失、不合法输出、reader-only。
2. Spec first：先更新 `docs/spec.md` 和必要的 `spec/ir-v1.schema.json`。
3. TDD：先写 IR、编译 XML/包、外部 XLSX reverse 测试，跑红灯后实现。
4. 正反向共用单位、枚举和语义 helper；范围/关系结构必须按 ID 配对重建。
5. 更新 README 功能矩阵和公开 fixture；不得引用私有路径或素材。
6. 本地质量门：fmt、`cargo test --all-targets --all-features --locked`、clippy、
   `git diff --check`；XLSX 包用 .NET `DocumentFormat.OpenXml` 做集成合法性验证。
7. 远端 GPU 验证只同步源码并增量构建；暂不可达时记录准确错误，继续本地工作。
8. 正反向、规范与验证完整的功能批次独立提交。推送后检查 GitHub Actions CI
   真实运行结果；CI 成功后使用 cargo-dist 计划并创建 GitHub Release，再用
   `gh` 确认 Release 和跨平台资产真实存在。标签存在不等于 Release 成功。
   Release/CI 流水线运行时可以继续下一个常用功能，但必须如实记录状态；若仓库
   remote 或 GitHub 仓库未配置，停止发布并向用户说明所需信息，不猜测或创建仓库。
9. 功能优先级以常见工作流和可无损正反向闭环为准。低频、上游 writer 缺失、输出
   不合法或难以无损逆向的能力明确写入 `docs/roadmap.md`，不为追求一次性全覆盖
   阻塞常用功能。

## 镜像与安全

- Cargo/crates.io 优先 `.cargo/config.toml` 的 rsproxy；有限重试，不因一次超时切回海外源。
- 保留用户已有工作；提交前检查工作树与是否意外包含真实业务 XLSX。
