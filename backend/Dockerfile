# ---- Build/runtime image for the Beacon backend ----
FROM node:22-slim

# Create app dir
WORKDIR /app

# Install deps first for better layer caching
COPY package.json ./
# Use npm install (no lockfile committed in this scaffold). For reproducible
# builds, commit package-lock.json and switch to: RUN npm ci --omit=dev
RUN npm install --omit=dev && npm cache clean --force

# Copy source
COPY src ./src

# The encrypted token store lives here; mount a volume to persist it.
RUN mkdir -p /app/.data && chown -R node:node /app
VOLUME ["/app/.data"]

ENV NODE_ENV=production
ENV PORT=4000
EXPOSE 4000

# Drop root
USER node

# Basic healthcheck hits the /health endpoint
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD node -e "fetch('http://localhost:'+(process.env.PORT||4000)+'/health').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"

CMD ["node", "src/server.js"]
