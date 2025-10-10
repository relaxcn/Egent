/**
 * TypeScript 类型声明：支持导入 Markdown 文件
 */

declare module "*.md" {
  const content: string;
  export default content;
}
