import streamlit as st
import streamlit.components.v1 as components


def run_welcome():

    components.html("""
<!DOCTYPE html>
<html>
<head>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@300;400;600&family=Inter:wght@300;400;500&display=swap" rel="stylesheet">
<style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
        background: transparent;
        overflow: hidden;
    }

    .scene {
        position: relative;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        min-height: 380px;
        padding: 3rem 2rem;
        overflow: hidden;
    }

    /* Single deep ambient glow - very subtle, centred */
    .deep-glow {
        position: absolute;
        width: 500px;
        height: 260px;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: radial-gradient(ellipse at center,
            rgba(80, 100, 220, 0.07) 0%,
            rgba(50, 70, 180, 0.04) 40%,
            transparent 70%);
        pointer-events: none;
        filter: blur(18px);
        animation: glowBreathe 9s ease-in-out infinite alternate;
    }

    /* Star canvas sits behind everything */
    #starCanvas {
        position: absolute;
        inset: 0;
        width: 100%;
        height: 100%;
        pointer-events: none;
    }

    /* Top rule line */
    .rule-top {
        position: relative;
        z-index: 2;
        width: 0;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(99,118,255,0.45), rgba(6,182,212,0.45), transparent);
        margin-bottom: 2.5rem;
        animation: expandRule 1.4s cubic-bezier(0.16, 1, 0.3, 1) forwards 0.3s;
    }

    /* Badge */
    .badge {
        position: relative;
        z-index: 2;
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        padding: 0.32rem 1rem;
        border: 1px solid rgba(99, 118, 255, 0.22);
        border-radius: 100px;
        background: rgba(99, 118, 255, 0.05);
        margin-bottom: 1.75rem;
        opacity: 0;
        animation: fadeUp 0.8s cubic-bezier(0.16, 1, 0.3, 1) forwards 0.6s;
    }
    .badge-dot {
        width: 5px; height: 5px;
        border-radius: 50%;
        background: #6376FF;
        animation: badgePulse 2.5s ease-in-out infinite;
    }
    .badge-text {
        font-family: 'Inter', sans-serif;
        font-size: 0.63rem;
        font-weight: 500;
        letter-spacing: 0.18em;
        text-transform: uppercase;
        color: rgba(99, 118, 255, 0.9);
    }

    /* Main title */
    .title-wrap {
        position: relative;
        z-index: 2;
        margin-bottom: 1.4rem;
        opacity: 0;
        animation: fadeUp 1s cubic-bezier(0.16, 1, 0.3, 1) forwards 0.9s;
    }
    .title {
        font-family: 'Cormorant Garamond', Georgia, serif;
        font-size: clamp(2.4rem, 6vw, 3.5rem);
        font-weight: 300;
        letter-spacing: 0.01em;
        line-height: 1.15;
        color: #E4E8F5;
    }
    .title em {
        font-style: italic;
        font-weight: 600;
        background: linear-gradient(120deg, #9BAEFF 0%, #C8CFFF 45%, #7DD6EA 100%);
        background-size: 220% auto;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        animation: shimmer 6s linear 2s infinite;
    }

    /* Subtitle */
    .subtitle {
        position: relative;
        z-index: 2;
        font-family: 'Inter', sans-serif;
        font-size: 0.88rem;
        font-weight: 300;
        color: rgba(130, 142, 172, 0.8);
        letter-spacing: 0.05em;
        line-height: 1.9;
        text-align: center;
        opacity: 0;
        animation: fadeUp 0.9s cubic-bezier(0.16, 1, 0.3, 1) forwards 1.2s;
    }

    /* Cursor */
    .cursor {
        display: inline-block;
        width: 1.5px;
        height: 0.85em;
        background: linear-gradient(180deg, #6376FF, #06B6D4);
        margin-left: 2px;
        vertical-align: middle;
        border-radius: 1px;
        animation: blink 1.1s ease-in-out infinite;
    }

    /* Bottom accent */
    .bottom-accent {
        position: relative;
        z-index: 2;
        display: flex;
        align-items: center;
        gap: 0.75rem;
        margin-top: 2.2rem;
        opacity: 0;
        animation: fadeUp 0.8s cubic-bezier(0.16, 1, 0.3, 1) forwards 1.5s;
    }
    .accent-line {
        width: 0;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(99,118,255,0.4));
        animation: expandAccent 1.1s cubic-bezier(0.16, 1, 0.3, 1) forwards 1.7s;
    }
    .accent-line.right {
        background: linear-gradient(90deg, rgba(6,182,212,0.4), transparent);
    }
    .accent-diamond {
        width: 4px; height: 4px;
        background: linear-gradient(135deg, #6376FF, #06B6D4);
        transform: rotate(45deg);
        border-radius: 0.5px;
        animation: diamondGlow 4s ease-in-out infinite alternate;
    }

    /* Bottom rule */
    .rule-bottom {
        position: relative;
        z-index: 2;
        width: 0;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(6,182,212,0.35), transparent);
        margin-top: 2.5rem;
        animation: expandRule 1.4s cubic-bezier(0.16, 1, 0.3, 1) forwards 1.9s;
    }

    /* Keyframes */
    @keyframes fadeUp {
        from { opacity: 0; transform: translateY(18px); }
        to   { opacity: 1; transform: translateY(0); }
    }
    @keyframes expandRule {
        from { width: 0; }
        to   { width: 260px; }
    }
    @keyframes expandAccent {
        from { width: 0; }
        to   { width: 56px; }
    }
    @keyframes shimmer {
        0%   { background-position: 0% center; }
        100% { background-position: 220% center; }
    }
    @keyframes blink {
        0%, 100% { opacity: 1; }
        45%, 55% { opacity: 0; }
    }
    @keyframes badgePulse {
        0%, 100% { opacity: 1; transform: scale(1); box-shadow: 0 0 0 0 rgba(99,118,255,0.4); }
        50%       { opacity: 0.6; transform: scale(0.85); box-shadow: 0 0 0 3px rgba(99,118,255,0); }
    }
    @keyframes glowBreathe {
        from { opacity: 0.7; transform: translate(-50%, -50%) scale(1); }
        to   { opacity: 1;   transform: translate(-50%, -50%) scale(1.08); }
    }
    @keyframes diamondGlow {
        from { opacity: 0.6; transform: rotate(45deg) scale(1); }
        to   { opacity: 1;   transform: rotate(45deg) scale(1.4); }
    }
</style>
</head>
<body>
<div class="scene">

    <div class="deep-glow"></div>
    <canvas id="starCanvas"></canvas>

    <div class="rule-top"></div>

    <div class="badge">
        <span class="badge-dot"></span>
        <span class="badge-text">Pibit.ai &nbsp;·&nbsp; Enterprise Platform</span>
    </div>

    <div class="title-wrap">
        <h1 class="title">Welcome to Insight Board</h1>
    </div>

    <p class="subtitle">
        Non-lossrun insights are here<span class="cursor"></span>
    </p>

    <div class="bottom-accent">
        <div class="accent-line"></div>
        <div class="accent-diamond"></div>
        <div class="accent-line right"></div>
    </div>

    <div class="rule-bottom"></div>

</div>

<script>
    const canvas = document.getElementById('starCanvas');
    const ctx = canvas.getContext('2d');

    function resize() {
        canvas.width  = canvas.offsetWidth;
        canvas.height = canvas.offsetHeight;
    }
    resize();

    // Build star field — three tiers: tiny dust, mid, rare bright
    const stars = [];
    const W = canvas.width  || 700;
    const H = canvas.height || 420;

    function buildStars() {
        stars.length = 0;

        // Tier 1: fine dust — many, barely visible
        for (let i = 0; i < 120; i++) {
            stars.push({
                x: Math.random() * W,
                y: Math.random() * H,
                r: Math.random() * 0.6 + 0.2,
                baseAlpha: Math.random() * 0.18 + 0.06,
                alpha: 0,
                twinkleSpeed: Math.random() * 0.004 + 0.001,
                twinkleOffset: Math.random() * Math.PI * 2,
                tier: 1
            });
        }

        // Tier 2: mid stars — fewer, slightly more visible
        for (let i = 0; i < 40; i++) {
            stars.push({
                x: Math.random() * W,
                y: Math.random() * H,
                r: Math.random() * 0.8 + 0.5,
                baseAlpha: Math.random() * 0.22 + 0.1,
                alpha: 0,
                twinkleSpeed: Math.random() * 0.003 + 0.001,
                twinkleOffset: Math.random() * Math.PI * 2,
                tier: 2
            });
        }

        // Tier 3: rare accent stars — very few, soft glow
        for (let i = 0; i < 8; i++) {
            stars.push({
                x: Math.random() * W,
                y: Math.random() * H,
                r: Math.random() * 1.0 + 0.9,
                baseAlpha: Math.random() * 0.28 + 0.12,
                alpha: 0,
                twinkleSpeed: Math.random() * 0.002 + 0.0005,
                twinkleOffset: Math.random() * Math.PI * 2,
                tier: 3
            });
        }
    }

    buildStars();

    let t = 0;
    function draw() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        t += 0.016;

        for (const s of stars) {
            // Gentle sinusoidal twinkle
            const alpha = s.baseAlpha * (0.6 + 0.4 * Math.sin(t * s.twinkleSpeed * 60 + s.twinkleOffset));

            if (s.tier === 3) {
                // Soft glow for accent stars
                const grd = ctx.createRadialGradient(s.x, s.y, 0, s.x, s.y, s.r * 3.5);
                grd.addColorStop(0, `rgba(190, 200, 255, ${alpha})`);
                grd.addColorStop(1, `rgba(190, 200, 255, 0)`);
                ctx.beginPath();
                ctx.arc(s.x, s.y, s.r * 3.5, 0, Math.PI * 2);
                ctx.fillStyle = grd;
                ctx.fill();
            }

            // Core dot
            ctx.beginPath();
            ctx.arc(s.x, s.y, s.r, 0, Math.PI * 2);
            const white = s.tier === 1 ? '210, 215, 240' : '200, 210, 255';
            ctx.fillStyle = `rgba(${white}, ${alpha})`;
            ctx.fill();
        }

        requestAnimationFrame(draw);
    }

    draw();
</script>
</body>
</html>
""", height=420)