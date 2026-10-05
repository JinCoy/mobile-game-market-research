import {defineConfig} from '@playwright/test';
const externalUrl = process.env.PLAYFIELD_TEST_URL;
export default defineConfig({testDir:'./tests',use:{baseURL:externalUrl || 'http://127.0.0.1:3000/',headless:true,channel:'chrome'},workers:1,reporter:'list',webServer:externalUrl ? undefined : {command:'npm run start',url:'http://127.0.0.1:3000',reuseExistingServer:true,timeout:60000}});
