#version 440

// Пыль растворения («танос»), один слой из трёх. Каждый слой — своя
// семья пылинок: своя скорость, размер и доля ячеек.
//
// Идея из Telegram Desktop (ui/effects/thanos_effect): случайность не
// хранят — вычисляют хешем от координаты и зерна, цвет пылинка берёт
// из самого снимка карточки. Пылинка здесь — мягкая капля, чуть
// вытянутая по полёту и сжимающаяся к концу жизни (как в их рендерере):
// не пиксельный шум, а заметная пыль.
//
// Устройство: сцена разбита на сетку ячеек (8 px); в ячейке до двух
// пылинок семьи. Пылинка «живёт» в ячейке, а точку карточки, откуда
// она летит, находим обратным путём: позиция минус смещение за время
// после пробуждения (волна справа налево — как в карточном варианте).
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
    float uSeed;      // зерно: каждый взрыв новый (у каждого слоя своё)
    vec2 uQuad;       // размер слоя, px
    vec4 uCard;       // карточка в координатах слоя: x, y, w, h
    vec4 uMotion;     // vx, vy — скорость семьи, px/с; zw — снос, px/с
    vec4 uClass;      // x — сколько пылинок на ячейку (может быть > 1),
                      // y — тяжесть, px/с², z — базовый радиус, px,
                      // w — разброс радиуса
    vec2 uLife;       // время жизни: база и разброс, с
};

float hash1(vec2 p) {
    return fract(sin(dot(p, vec2(127.1, 311.7)) + uSeed * 17.13) * 43758.5453);
}

void main() {
    vec2 pos = qt_TexCoord0 * uQuad;

    // быстрый отсекатель: дальше запаса от карточки пыли не бывает
    if (pos.x < uCard.x - 30.0 || pos.x > uCard.x + uCard.z + 160.0 ||
        pos.y < uCard.y - 90.0 || pos.y > uCard.y + uCard.w + 30.0) {
        discard;
    }

    const float cell = 8.0;      // ячейка сетки пылинок, px
    const float kDecay = 7.7;    // затухание начальной скорости, 1/с
    vec2 baseCell = floor(pos / cell);

    vec3 accRGB = vec3(0.0);
    float accA = 0.0;

    // 5x5 соседних ячеек: капля радиусом до ~5 px с вытяжкой и сдвигом
    // достаёт на ~11 px, это кольцо обязано быть проверенным
    for (int j = -2; j <= 2; j++) {
        for (int i = -2; i <= 2; i++) {
            vec2 c = baseCell + vec2(float(i), float(j));
            for (int s = 0; s < 2; s++) {
                vec2 cs = c + vec2(float(s) * 0.37, float(s) * 0.71);
                float gate = clamp(uClass.x - float(s), 0.0, 1.0);
                if (hash1(cs) > gate) continue;   // пылинки нет

                float hSpd  = hash1(cs + vec2(7.7, 3.1));
                float hAng  = hash1(cs + vec2(5.3, 9.2));
                float hWob  = hash1(cs + vec2(23.0, 5.0));
                float hLife = hash1(cs + vec2(13.9, 27.4));
                float hSize = hash1(cs + vec2(3.3, 17.2));
                float hJx   = hash1(cs + vec2(11.7, 21.3));
                float hJy   = hash1(cs + vec2(19.1, 8.7));

                // позиция пылинки: центр ячейки со случайным сдвигом,
                // чтобы пыль не строилась в видимую решётку
                vec2 pPos = (c + vec2(0.5)) * cell + (vec2(hJx, hJy) - vec2(0.5)) * 6.0;

                // волна: правый край карточки просыпается первым,
                // момент пробуждения — обратная кривая торможения OutQuad
                float fx = clamp((pPos.x - uCard.x) / uCard.z, 0.0, 1.0);
                float th = (1.0 - fx) * 0.7 + hWob * 0.3;
                float tau = uT - 0.5 * (1.0 - sqrt(max(0.0, 1.0 - th)));
                if (tau <= 0.0) continue;

                // путь от карточки к ячейке: рывок с затуханием + снос + тяжесть
                float sp = 0.6 + hSpd * 0.8;
                float ang = (hAng - 0.5) * 0.7;
                float ca = cos(ang);
                float sa = sin(ang);
                vec2 v0 = vec2(ca * uMotion.x - sa * uMotion.y,
                               sa * uMotion.x + ca * uMotion.y) * sp;
                float damp = 1.0 - exp(-kDecay * tau);
                vec2 off = v0 / kDecay * damp + uMotion.zw * tau;
                off.y += 0.5 * uClass.y * tau * tau;
                off.x += sin(tau * 9.0 + hAng * 6.2832) * 2.5;
                off.y += cos(tau * 11.0 + hSpd * 6.2832) * 2.5;

                // точка карточки, откуда летит эта пылинка
                vec2 origin = pPos - off;
                vec2 uv = (origin - uCard.xy) / uCard.zw;
                if (uv.x < 0.0 || uv.x > 1.0 || uv.y < 0.0 || uv.y > 1.0) continue;

                vec4 col = texture(src, uv);
                if (col.a <= 0.04) continue;

                // жизнь → гашение и сжатие: пылинка тает
                float life = uLife.x + hLife * uLife.y;
                float a = clamp(1.0 - tau / life, 0.0, 1.0);
                if (a <= 0.0) continue;

                // мягкая капля, вытянутая по полёту, к концу сжимается
                vec2 vel = v0 * exp(-kDecay * tau) + uMotion.zw;
                float vl = length(vel);
                vec2 dir = vl > 0.001 ? vel / vl : vec2(1.0, 0.0);
                float r = (uClass.z + hSize * uClass.w) * (0.35 + 0.65 * a);
                vec2 q = pos - pPos;
                q = vec2(dot(q, dir), dot(q, vec2(-dir.y, dir.x)));
                q /= vec2(r * 1.6, r * 0.9);
                float m = 1.0 - smoothstep(0.35, 1.0, length(q));
                if (m <= 0.0) continue;

                // поверх уже накопленного (premultiplied source-over)
                float w = a * m * qt_Opacity;
                accRGB += (1.0 - accA) * col.rgb * w;
                accA += (1.0 - accA) * col.a * w;
            }
        }
    }

    if (accA <= 0.004) discard;
    fragColor = vec4(accRGB, accA);
}
