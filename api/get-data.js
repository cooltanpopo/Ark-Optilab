import { Redis } from '@upstash/redis'

// Check if we have the environment variables properly set by Vercel Upstash Integration
const url = process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL || '';
const token = process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN || '';

if (!url || !token) {
    console.warn("Redis credentials missing. Please ensure Upstash Redis is linked to this Vercel project.");
}

const redis = new Redis({
    url,
    token,
})

export default async function handler(req, res) {
    if (req.method !== 'GET') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const { key } = req.query;
    if (!key) {
        return res.status(400).json({ error: 'Key is required' });
    }

    try {
        const value = await redis.get(key);
        // Redis returns null if key doesn't exist
        res.status(200).json({ value });
    } catch (error) {
        console.error(`Error fetching key "${key}" from Redis:`, error);
        res.status(500).json({ error: 'Failed to fetch data' });
    }
}
