import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base: 깃허브 저장소 이름으로 바꿔주세요 (예: 저장소가 ffe-ose-order-manager 라면 아래 값 그대로 사용)
// 예: https://<username>.github.io/<repo-name>/ 형태로 배포되므로 base는 "/<repo-name>/" 이어야 합니다.
export default defineConfig({
  plugins: [react()],
  base: "/ffe-ose-order-manager/",
});
