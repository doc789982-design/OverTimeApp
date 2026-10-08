#version 440

// Пыль растворения («танос»), один слой из трёх.
//
// Идея из Telegram Desktop (ui/effects/thanos_effect): случайность не хранят —
// её вычисляют на лету хешем от координаты и зерна, а цвет пылинка берёт
// из самого снимка карточки. Здесь каждый пиксель слоя сам решает, чья он
// пылинка: определяет свой класс и плотность, скорость с разбросом,
// момент пробуждения (волна справа налево — как в карточном варианте),
// и подбирает исходный пиксель карточки, откуда он летит.
//
// Выход — premultiplied (требование ShaderEffect), сэмплер на binding 1,
// uniform-блок обязан начинаться с qt_Matrix и qt_Opacity (дока Qt:
// fragment-only шейдеры делят буфер со встроенным вершинным).

layout(location = 0) in vec2 qt_TexCoord0;
layout(location = 0) out vec4 fragColor;

layout(binding = 1) uniform sampler2D src;

layout(std140, binding = 0) uniform buf {
    mat4 qt_Matrix;
    float qt_Opacity;
    float uT;         // секунды с начала эффекта
    float uSeed;      // зерно: каждый взрыв новый
    vec2 uQuad;       // размер слоя, px
    vec4 uCard;       // карточка в координатах слоя: x, y, w, h
    vec4 uMotion;     // vx, vy — скорость класса, px/с; zw — постоянный снос, px/с
    vec4 uClass;      // lo, hi — чья это пыль; density — плотность; w — тяжесть, px/с²
    vec2 uLife;       // время жизни: база и разброс, с
};

float hash1(vec2 p) {
    return fract(sin(dot(p, vec2(127.1, 311.7)) + uSeed * 17.13) * 43758.5453);
}

void main() {
    vec2 pos = qt_TexCoord0 * uQuad;

    // чей этот пиксель: класс (треть всего хеша на слой) и плотность
    float h1 = hash1(pos);
    if (h1 < uClass.x || h1 >= uClass.y) discard;
    if (hash1(pos + vec2(31.7, 11.3)) > uClass.z) discard;

    // скорость пылинки: база класса, повёрнутая и промасштабированная хешем
    float hs = hash1(pos + vec2(7.7, 3.1));
    float ha = hash1(pos + vec2(5.3, 9.2));
    float sp = 0.6 + hs * 0.8;
    float ang = (ha - 0.5) * 0.7;
    float ca = cos(ang);
    float sa = sin(ang);
    vec2 v = vec2(ca * uMotion.x - sa * uMotion.y,
                  sa * uMotion.x + ca * uMotion.y) * sp;

    // волна: правый край карточки просыпается первым (порог как в карточном
    // варианте), момент пробуждения — обратная кривая торможения OutQuad
    float fx = (pos.x - uCard.x) / uCard.z;
    float th = (1.0 - clamp(fx, 0.0, 1.0)) * 0.7 + hash1(pos + vec2(23.0, 5.0)) * 0.3;
    float wake = 0.5 * (1.0 - sqrt(max(0.0, 1.0 - th)));
    float tau = uT - wake;
    if (tau <= 0.0) discard;

    // путь: рывок с затуханием + постоянный снос + лёгкая тяжесть + покачивание
    float k = 7.7; // затухание начальной скорости, 1/с (0.88 в кадр)
    vec2 off = v / k * (1.0 - exp(-k * tau)) + uMotion.zw * tau;
    off.y += 0.5 * uClass.w * tau * tau;
    off.x += sin(tau * 9.0 + ha * 6.2832) * 2.5;
    off.y += cos(tau * 11.0 + hs * 6.2832) * 2.5;

    // исходный пиксель карточки: откуда эта пылинка летит
    vec2 origin = pos - off;
    vec2 uv = (origin - uCard.xy) / uCard.zw;
    if (uv.x < 0.0 || uv.x > 1.0 || uv.y < 0.0 || uv.y > 1.0) discard;

    vec4 c = texture(src, uv);
    if (c.a <= 0.04) discard;

    // время жизни → гашение (1.0–1.4 «единиц» карточного варианта = 1.39–1.94 с)
    float life = uLife.x + hash1(pos + vec2(13.9, 27.4)) * uLife.y;
    float a = clamp(1.0 - tau / life, 0.0, 1.0) * qt_Opacity;

    fragColor = vec4(c.rgb * c.a * a, c.a * a);
}
