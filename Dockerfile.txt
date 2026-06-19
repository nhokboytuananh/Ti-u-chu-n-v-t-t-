# Sử dụng image Node.js thu gọn
FROM node:20-alpine AS builder

# Thư mục làm việc trong container
WORKDIR /app

# Copy package.json và cài đặt dependencies
COPY package*.json ./
RUN npm install

# Copy toàn bộ mã nguồn
COPY . .

# Build ứng dụng (Vite Frontend + Express Backend)
RUN npm run build

# Chuyển sang giai đoạn Production
FROM node:20-alpine

WORKDIR /app

# Copy package.json và chỉ cài đặt các dependencies cần cho chạy thật (loại bỏ devDependencies)
COPY package*.json ./
RUN npm ci --omit=dev

# Copy thư mục dist đã biên dịch từ bước builder
COPY --from=builder /app/dist ./dist

# Mở cổng 3000
EXPOSE 3000

# Gán biến môi trường
ENV NODE_ENV=production
ENV PORT=3000

# Khởi chạy server
CMD ["npm", "start"]
