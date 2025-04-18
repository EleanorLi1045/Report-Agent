import json
import os
import sys
from openai import OpenAI

def load_openai_key(key_file="OpenAIKey.txt"):
    with open(key_file, "r") as f:
        return f.read().strip()

def get_ring_progress_from_status(status_tweet_for_ai):
    client = OpenAI(api_key=load_openai_key())

    examples = """
Example 1:
Status: Work is planned for this semester but has not started
Rings: {"ring_0": 0, "ring_1": 0, "ring_2": 0, "ring_3": 0, "ring_4": 0}

Example 2:
Status: Feature flag added to run experiment in R0
Rings: {"ring_0": 100, "ring_1": 0, "ring_2": 0, "ring_3": 0, "ring_4": 0}

Example 3:
Status: AL changes completed. Integrating with TMP now. Telemetry remaining.
Rings: {"ring_0": -1, "ring_1": -1, "ring_2": -1, "ring_3": -1, "ring_4": -1}

Example 4:
Status: 04/02 - Client rolled out to R4 100%, service rolled out to WW 50%
Rings: {"ring_0": 100, "ring_1": 100, "ring_2": 100, "ring_3": 100, "ring_4": 50}

Example 5:
Status: Rolled out to R2 100%, R3 rollout pending
Rings: {"ring_0": 100, "ring_1": 100, "ring_2": 100, "ring_3": 0, "ring_4": 0}

Example 6:
Status: R0 FC testing underway -> AI Alpha SG by 04/04, R0 by 04/18 and today's date is 4/3
Rings: {"ring_0": 0, "ring_1": 0, "ring_2": 0, "ring_3": 0, "ring_4": 0}

Example 7:
Status: Released 100% in desired ring
Rings: {"ring_0": 100, "ring_1": 100, "ring_2": 100, "ring_3": 100, "ring_4": 100}
"""

    prompt = f"""
You are an AI assistant that classifies software deployment progress across 5 deployment rings (ring_0 to ring_4) based on the deployment status text.

Rules:
1. If a release (e.g. R0, R1) is mentioned with a **future date**, mark it as 0%.
2. If any ring has progress (like R2 = 50%), all lower rings (R0, R1) **must be 100%**.
3. Each ring must be <= 100 and >= the previous ring.
4. Respond only in JSON format.

{examples}

Now analyze this:

Status: {status_tweet_for_ai}

Rings:
"""

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are a helpful assistant that classifies deployment status into rollout rings."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.2,
        max_tokens=150
    )

    reply = response.choices[0].message.content.strip()
    try:
        return json.loads(reply)
    except json.JSONDecodeError:
        return ""