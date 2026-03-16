import { Redis } from '@upstash/redis'

const url = process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL || '';
const token = process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN || '';

if (!url || !token) {
    console.warn("Redis credentials missing. Please ensure Upstash Redis is linked to this Vercel project.");
}

const redis = new Redis({
    url,
    token,
})

// Increase the payload size limit up to 4MB for Vercel Serverless Functions
// Note: 4.5MB is Vercel's standard limit for function payloads
export const config = {
    api: {
        bodyParser: {
            sizeLimit: '4mb',
        },
    },
}

export default async function handler(req, res) {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const { key, value } = req.body;

    if (!key) {
        return res.status(400).json({ error: 'Key is required' });
    }

    try {
        // Save to Upstash Redis
        await redis.set(key, value);
        res.status(200).json({ success: true });
    } catch (error) {
        console.error(`Error saving key "${key}" to Redis:`, error);
        res.status(500).json({ error: 'Failed to save data' });
    }
}
