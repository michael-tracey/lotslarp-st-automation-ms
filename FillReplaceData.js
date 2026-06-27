/**
 * @OnlyCurrentDoc
 *
 * Sample data for the 'fill-replace-data' sheet, used by Initialise Project to
 * seed a fresh spreadsheet so the "Fill cell with Feed/Herd/Patrol/Discipline
 * Data" tools work out of the box.
 *
 * How the fill tool reads this sheet (see fillCellWithData_ in Actions.js):
 *  - The 4 base columns (Herd, Feed, Patrol, Discipline) hold template strings.
 *    One is picked at random for the chosen action type.
 *  - Any [placeholder] in that template is replaced with a random non-empty
 *    value from the column whose header matches the placeholder name.
 *  - Each column is an INDEPENDENT pool: empty cells are filtered out, so the
 *    columns do not need to be the same length and row alignment is irrelevant.
 *
 * This is a representative sample drawn from the production sheet — enough that
 * every placeholder used by the templates has values to draw from. Storytellers
 * can expand any column directly in the sheet afterwards.
 */
function getFillReplaceSampleData_() {
    const headers = [
        'Herd', 'Feed', 'Patrol', 'Discipline', 'person', 'metaphor', 'time', 'noise',
        'business', 'smell', 'light', 'animal', 'observed_detail', 'weather_element',
        'wolfsign', 'disc_attain', 'disc_beast_1', 'disc_beast_2', 'disc_feeling',
        'old-patrols-dumping-ground', 'wolfsign_boring', 'metaphor_unbidden',
        'thought_theme', 'metaphor_lingering'
    ];

    // Each key is a column; values are its independent pool of options.
    const pools = {
        'Herd': [
            "You meet up with your herd as usual. As your fangs sink into their skin, a thought comes to you unbidden, like [metaphor_unbidden]. [thought_theme] You stay quiet throughout, and when it's over, you depart from your herd, but the thought lingers like [metaphor_lingering]."
        ],
        'Feed': [
            "As you feed on [person], a thought occurs to you. [thought_theme] You finish your feeding, and you leave your victim where you found them, but the thought cannot be abandoned so easily. It stays with you like [metaphor_lingering]."
        ],
        'Patrol': [
            "[PATROL SUCCESS] - You wander through the immaculate and impossibly green lawns of an upper middle class not-quite-gated community. From behind a scultped topiary bush, you look out on the cookie cutter houses and beige-curtained windows. Near the swimming pool, you see [person] , the cold moon emphasizing the shadows of their face. The scene is so bland, so excrutiatingly *normal*, it's like a [metaphor]. You continue your patrol, feeling fully your inhuman nature. You aren't made for scenes like this, if you ever were. Now you're just a discordant note in their tinny sitcom soundtrack.",
            "[PATROL SUCCESS] You wind your way through back alleys and one-way streets, broken bricks kicked off to the side and mouse-bitten advertisements littering the ground beneath your feet. You let the wind guide you through the night, apathetically listening to the sounds of [noise], until your eyes catch on a [light]. Curiosity gets the better of you, but as you start to head towards it you lose sight of it, and soon after any interest. The night is quiet, almost boring, like [metaphor], leaving you with nothing in particular to report back.",
            "[PATROL SUCCESS] - You stop to rest in a clump of trees edging a business park, listening to the quiet sounds of mortals in the distance, the biting cold of proper winter settling on your skin like a pall. The path through the buildings and trees is shrouded in darkness, just like your future. A sudden noise from near the street, [noise], startles you. There's another sound from just a couple trees down, probably a [animal], and you let yourself relax, remembering your place as an apex predator, but then you notice something: [wolfsign_boring]. The beast inside you groans disconsolately. Everything is flat and safe and boring. You wonder when this oasis of calm is going to end.",
            "[PATROL SUCCESS] - The long, winding back road is silent except for the sound of [noise] from somewhere a street over. You hear your footsteps a little too loundly as you make your way in that direction. From just ahead you suddenly catch sight of [person], and the smell of [smell] wafts over you discordantly. You come to a crossroads, and jump back to the shadows as your eyes catch on something: [wolfsign_boring]. You watch for a long moment before letting your guard down, shrugging, and heading on down the road. There's still nothing to really see, just like all the other nights recently. The future, it seems, is beige.",
            "[PATROL SUCCESS] You come to rest on a small bench, having had nothing grab your attention so far in the night. You hear something odd from a [business] across the street, but looking over, you see [person] and somehow it all makes sense. With nothing more to see, you make your way along, getting lost in the repetitiveness of the night, nearly abandoning your post, when you finally catch something, freezing as it starts to process… and then as quickly as it caught your eye, you realize it’s simply [wolfsign_boring]. It means nothing, now, but you don’t forget why you’re out here until the morning comes.",
            "[PATROL SUCCESS] - You move through a cemetery, listening to the sound of a funerary tent flapping uselessly in the incessant wind. The endless, boring nothingness here is like a [metaphor], a deflated reminder of your own, pointless existence. The weather seems to be changing, an incoming [weather_element]. Your patrol feels aimless and indescribably boring until you suddenly see [wolfsign_boring]. You watch for a moment, lost in the pointlessness of it all. And then you walk out through the cemetery gates, keeping half an eye out around you but knowing that nothing is going to happen.",
            "[PATROL SUCCESS] - You make your way around yet another new construction in your territory: cheaply made and undoubtedly overpriced condos, decaying in the wintry night even before the cheap vinyl siding has been attached. There's a discordant smell of [smell], and a persistent sound of [noise]. You see [person], just going about their business with no thought to you, or to anything in the world outside themself, as far as you can tell. You wander along, noting the piles of construction debris and trash, barrels of some kind of solvent slowly leaking into the earth unheeded. It's all....pointless; [metaphor]. Nothing is going to happen tonight. Nothing seems to happen anywhere, anymore.",
            "[PATROL SUCCESS] You wander along the outskirts of the territory, your eye catching on [observed_detail]. You’ve seen the same sort of thing dozens of times before, and at [time], you think that it’s nothing, really. It doesn’t point to any exciting activity, isn’t anything worth following, and much like the rest of your night, it’s like [metaphor]. Just another night leaving you wondering what the point of this even is.",
            "[Patrol Success] - The [weather_element] descends on the city, disrupting the sound of [noise]. You pause just at the corner of a closed [business], letting a group of mortals pass, aware that it is [time]. The last is [person], and you wonder what they would do if you went up and actually spoke to them. Tried to describe what you are and what your unlife is actually like. That you literally eat people like them for breakfast. Your beast rouses, but you let them pass, sinking back into your facade of normal, until your eyes land on [wolfsign_boring]. You stare at it fixedly for a moment, before letting yourself shrug and move along. It means nothing. None of it actually *means* anything, anymore.",
            "[PATROL SUCCESS] You move silently through your territory, like a ghost, watching for any sign of disturbance or trespassing. The sound of [noise] from nearby startles you, but you immediately make yourself relax, remembering these are safer nights, at least for the moment. Even probing the ghoul network for gossip, the best story you got back was [wolfsign_boring]. Looking around for anything more interesting to send back, you see [person]. A [animal] makes a quiet noise from the other direction. These things are unrelated. There's no narrative here, no...anything. You slouch your way back to your haven at [time]. All of this has the excitement of [metaphor]."
        ],
        'Discipline': [
            "**Success!**\nWhat will you learn this evening? Will you [disc_attain]? Only time will tell.\nYou focus your energy on your blood, and you let your beast come through for just a moment. Like [disc_beast_1] it arrives. You know that the discipline lies within you, that it is a property of your blood, but you are not yet aware of how to use it. No matter how hard you try, you cannot do it on your own, and then [disc_beast_2].\nWhen you're done you feel [disc_feeling] and you know that you can use this discipline at any time now. Whether you learned it on your own or with a teacher is not important, because the beast is the only teacher you ever really need."
        ],
        'person': [
            "a person dressed as Krampus, lounging on the curb near the entry to the Biltmore Estate. As you watch, they pull a joint from their pocket and light up, leaning back on their hands and staring at the sky while exhaling a massive cloud of smoke.",
            "a man running a knife along the outline of his fingers, tears trickling down his face",
            "a couple of teenagers, clearly having snuck out to be together, alternately giggling and making out",
            "someone dressed as a clown, doing the scarf trick to some beleagured onlookers, but all the scarves seem to be panties",
            "a person dressed as Darth Vader, wrapped in Christmas lights, walking drunkenly down the street calling out \"Ho ho ho\" in a deep baritone",
            "a young woman laden with far too many wrapped presents, trying to get from her car to somewhere farther down the street, leaving a trail of small gifts behind her",
            "a young person rollerskating in a parking lot, dressed in a Sonic costume, seemingly muttering \"Gotta go fast\" over and over",
            "a mime on stilts, looking down at the passersby and indicating that it cannot get to them through the glass box surrounding them",
            "a very tall woman walking out of a bar with a smirk on her face, her lipstick smeared and her hair mussed",
            "half a dozen people in various rainbow-striped clothing, singing a gay rights anthem and cavorting with each other as they wander to their next destination"
        ],
        'metaphor': [
            "a beige Honda Civic",
            "a deflated helium balloon draped over a branch",
            "a listless pair of shoes dangling from a power line",
            "a soggy sponge, left in the bottom of the bathtub",
            "akin to watching paint dry",
            "like eating cake someone forgot to put sugar in",
            "mold at the bottom of a glass",
            "a piece of glass between you and the world, making everything feel unreal"
        ],
        'time': [
            "the point the moon hits its zenith",
            "the hour when most mortals are fast asleep",
            "the darkest part of the night, when only owls and rats are still going about their business",
            "the point all the bars are closing and everyone is being shuffled out",
            "the young night, when there's still the sense that anything could happen",
            "the hour that even the highway is quiet",
            "the edge of night, that fragile, liminal boundary just before the first stirrings of the waking world intrude on the darkness"
        ],
        'noise': [
            "drone of a plane overhead",
            "tap tap of someone up late on their computer",
            "traffic, as people bustle to and fro in the night",
            "an old, crackly radio",
            "a clown horn",
            "someone sharpening knives",
            "a quiet conversation in a late night restaurant"
        ],
        'business': [
            "old bookstore",
            "craft brewery",
            "independent record store",
            "antique shop",
            "farm-to-table restaurant",
            "closed coffee shop",
            "library"
        ],
        'smell': [
            "decaying roses",
            "smoke coming from an all night barbeque joint",
            "vaguely burnt synthetics and overly pungent laundry detergent",
            "car exhaust and burnt rubber",
            "mildewing laundry",
            "rich chocolate coming from an all night chocolatier"
        ],
        'light': [
            "flickering security lamp",
            "buzzing neon sign",
            "warm patio string lights",
            "harsh fluorescent streetlights",
            "distant car headlights sweeping past",
            "blue bug zapper light behind the building"
        ],
        'animal': [
            "rat",
            "stray cat",
            "raccoon",
            "opossum",
            "skunk",
            "coyote"
        ],
        'observed_detail': [
            "a bent engagement ring, lying on the ground",
            "a single boot, left abandoned in the middle of the road",
            "a pile of pennies, left abandoned on the sidewalk",
            "a half-smoked joint, wrapped in what seems to be peppermint-striped paper"
        ],
        'weather_element': [
            "light drizzle misting the air, coating everything in a fine sheen, muffling sounds",
            "thick fog rolling in from the valleys, swallowing landmarks and creating a world of ghosts",
            "sudden gust of wind rattling signs, an invisible hand briefly disturbing the night's stillness",
            "distant rumble of thunder, the mountains muttering in their sleep, promising a storm",
            "bright moonlight illuminating everything, casting sharp shadows"
        ],
        'wolfsign': [
            "Two massively large dogs. One lunges at the other and the pair roll. The snapping of teeth and the growls seem to go unnoticed except by you. One of them gets free. It flees. The other chases after.",
            "A house on fire. No one's inside. Massive dogs...wolves? Maybe wolves...run circles around the house, and they howl at the moon. Out in the country, the fire department arrive too late for whatever burned inside.",
            "A pickup truck at a stop light. The bed is full of massively huge dogs with all sorts of visible injuries. The driver pulls their baseball cap down to cover their eyes. One yellow fang is clearly visible in the green light from the traffic signal as they pull away."
        ],
        'disc_attain': [
            "reach a height you've never before reached",
            "prove all the doubters wrong",
            "take your destiny in your own hands",
            "grasp the potential within you",
            "grow in a way you used to dream about"
        ],
        'disc_beast_1': [
            "a wave crashing on the side of a sinking ship",
            "a bullet ripping through the air",
            "a sound from a million years in the past",
            "a sunbeam",
            "the weight of centuries"
        ],
        'disc_beast_2': [
            "the beast does it for you",
            "you feel the beast take control of your blood for just a moment",
            "a surprising ease washes over you and you're in the passenger seat letting the beast take the wheel",
            "it's already done, because the beast did it for you when you weren't looking",
            "the movie of your success plays across your retinas, but your beast is the star, not you"
        ],
        'disc_feeling': [
            "drained",
            "accomplished",
            "smug with satisfaction",
            "stronger than you ever have",
            "like having a smoke and a nap"
        ],
        // Inactive "dumping ground" column — not referenced by the fill tool. Left empty.
        'old-patrols-dumping-ground': [],
        'wolfsign_boring': [
            "long-dried spraypaint, beginning to flake off the side of a building, which once read \"Treaty Broken\"",
            "old scratch marks on the wooden door of a store, once a reminder of the werewolves' persistent terror, now just an old scar",
            "a broken telephone pole the workers forgot to fix; a reminder of the terrifying power of the werewolves, now feeling more and more like a nightmare you can almost forget",
            "a street lined with lidless trash bins, a reminder of the werewolves' senseless havoc, now seemingly directed elsewhere"
        ],
        'metaphor_unbidden': [
            "an uninvited guest"
        ],
        'thought_theme': [
            "What if the wider Camarilla abandons Asheville, continuing the isolation of the domain but this time as a social denial rather than a geographic one?"
        ],
        'metaphor_lingering': [
            "an echo off in the distant hills sounding its persistent percussion"
        ]
    };

    // Build a rectangular grid (rows × headers). Each column is padded with
    // blanks to the length of the longest pool; the fill tool filters blanks.
    const maxLen = Math.max(...headers.map(h => (pools[h] || []).length));
    const rows = [];
    for (let r = 0; r < maxLen; r++) {
        rows.push(headers.map(h => {
            const pool = pools[h] || [];
            return r < pool.length ? pool[r] : '';
        }));
    }

    return { headers, rows };
}
