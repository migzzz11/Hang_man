Public Class frm_main

    'global variable declaration
    'word pool
    Dim wordPool() As String = {"abracadabra", "acappella", "acoustic", "acquiesce", "admiral", "aerobics", "agriculture", "amber", "ampere", "anchovy", "annuity", "annul", "antimony", "antique", "apparatus", "aparel", "aqueous", "arable", "army", "aroma", "arthritis", "arthropod", "askew", "asiant", "assess",
                               "babble", "backstitch", "bagel", "baggage", "barbecue", "barber", "battalion", "bawl", "bedazzle", "behoove", "bereave", "bicker", "bier", "bin", "bismuth", "bison", "blaspheme", "bleachers", "blossom", "blunder", "bogus", "bookkeeping", "boulevard", "broccoli", "briefcase",
                                "canary", "cancer", "cardiology", "cargo", "cavalry", "cedar", "chalice", "champagne", "cheese", "cheetah", "choker", "cholesterol", "circumcise", "circus", "clavicle", "claw", "coach", "coalesce", "coincidence", "coliseum", "comet", "commence", "compel", "compendium", "concatenate",
                                "dab", "daredevil", "daughter", "debacle", "debauch", "decennial", "deception", "decor", "deduct", "defiance", "defraud", "deign", "delicacy", "demitasse", "demolish", "deodorant", "deoxidize", "deprive", "deputation", "desecrate", "desolate", "deteriorate", "dethrone", "deviant", "dew",
                                "earphone", "earring", "eclair", "eclectic", "efficient", "effigy", "electrotype", "elegance", "embody", "embolden", "enable", "enact", "enigma", "enjoin", "entreat", "entrench", "epoch", "epoxy", "erratic", "erroneous", "estuary", "etch", "evergreen", "evoke", "exclaim",
                                "factory", "fallow", "falter", "farrago", "fastback", "feces", "federation", "fervor", "fescue", "fig", "fighter", "fingerprint", "fir", "flareup", "flashlight", "floodlight", "floor", "follicle", "folly", "forenoon", "foresight", "fortress", "fortune", "fraud", "fray",
                                "gabardine", "gamut", "gander", "gem", "gendarme", "germ", "germicide", "glamour", "gland", "gong", "gonorrhea", "grandparent", "granite", "grime", "grip", "guerrilla", "guise", "gyve", "guardian", "grouch", "grease", "gravy", "grace", "gown", "glue",
                                "haberdasher", "hair", "hake", "handbag", "hangar", "hansom", "happening", "harlot", "harmonica", "haulage", "haunch", "headstall", "health", "hearth", "heat", "helmet", "helve", "herd", "heritage", "highfalutin", "highness", "hirsute", "histamine", "holster", "homage",
                                "ibes", "icing", "icon", "iguana", "illness", "imitation", "imp", "impetus", "implement", "impression", "impulse", "income", "indecision", "independence", "index", "infamy", "infancy", "ingredient", "iniquity", "input", "inquest", "instance", "instinct", "intention", "intercom",
                                "jackal", "jawbreaker", "jazz", "jitterbug", "jive", "judgement", "judo", "junkie", "junta", "juxtapose", "jumbo", "july", "journal", "jonquil", "jettison", "jetsam", "jamb", "jalousie", "jeopardy", "joist", "juggernaut", "junction", "jury", "juncture", "jubilee",
                                "kabuki", "keel", "keepsake", "kilobit", "kilt", "knish", "knob", "kungfu", "kip", "kiosk", "keynote", "keyboard", "kamikaze", "kazoo", "keg", "kidney", "kin", "kiwi", "knout", "kook", "knife", "kiln", "kernel", "kapok", "karat",
                                "label", "lacrosse", "lactose", "landlord", "lane", "latchkey", "latex", "lawrencium", "lawsuit", "ledge", "leech", "lemur", "lens", "liar", "libel", "ligature", "lightning", "lineage", "lithium",
                                "macaroni", "magpie", "malaria", "malpractice", "manger", "manuscript", "marinade", "masacre", "matinee", "measles", "medley", "membrane", "merchandise", "metabolism", "mezzanine", "migraine", "minuet", "misdemeanor", "misogamy", "moisturize", "monastery", "monomania", "moratorium", "mosquito",
                                "nasturtium", "narcosis", "nationality", "nape", "nectar", "nausea", "naturalize", "navel", "nefarious", "negligent", "neoplasm", "nepotism", "neuritis", "newspaper", "nibble", "nightmare", "nitrogen", "nocturnal", "nonchalant", "nobility", "nostalgia", "novelty", "notarize", "notorious", "noxious",
                                "obelisk", "obituary", "obsolete", "obstinate", "occult", "octagon", "odyssey", "offspring", "ointment", "oligarchy", "omnipotent", "onslaught", "ooze", "opossum", "optometry", "orangutan", "orchestra", "oregano", "orifice", "osmosis", "ostracize", "outcast", "outrage", "ovary", "overdose",
                                "pageant", "pantomime", "parakeet", "parochial", "pathology", "pavillion", "penicillin", "peninsula", "permeate", "pessimism", "phenomenon", "phonograph", "pistachio", "placate", "plausible", "pneumonia", "pompous", "portfolio", "postmortem", "poverty", "preconceive", "prenatal", "primitive", "prodigal", "prologue",
                                "qiviut", "quadrangle", "quagmire", "quackery", "quarantine", "quaim", "quake", "quail", "queen", "quartz", "quarrel", "quartet", "queer", "quench", "query", "queue", "quinsy", "quirk", "quintet", "quill", "quotient", "quota", "quote", "quorum", "quoin",
                                "rabbit", "raccoon", "radish", "rampage", "ranci", "ransack", "rapacious", "rapture", "rascal", "raspberry", "ravage", "rebelion", "receptacle", "recreation", "recuperate", "rejuvenate", "repertoire", "reprimand", "resilient", "restaurant", "reverberate", "rhetoric", "rhinoceros",
                                "sacrifice", "sanatorium", "saffold", "scuttie", "semantics", "senile", "shadow", "shoulder", "shriek", "significance", "skein", "skeptic", "slaughter", "sleek", "sleigh", "sluggish", "smolder", "smudge", "snicker", "solemn", "solitaire", "sonnet", "squabble", "spontaneous", "stadium", "stomach", "subordinate",
                                "tabloid", "tantalize", "taxidermy", "telescope", "textile", "thaw", "theatrical", "therapy", "thieve", "thorax", "treshold", "throttle", "thunderbolt", "thyroid", "toboggan", "toffee", "tombstone", "tornado", "torpedo", "toxin", "traitor", "trauma", "tripod", "tremendous", "tumult", "tumor", "turquoise", "tycoon", "tyranny",
                                "ultimatum", "ultrasonic", "umbilical", "unanimous", "unconcious", "undertaker", "undying", "unfortunate", "unlucky", "unison", "unravel", "unreliable", "uniform", "upholster", "urge", "urgent", "uterus", "upgrade",
                                "vaccinate", "vandalism", "vapor", "vehicle", "vault", "venison", "vascular", "vesper", "viaduct", "vicarious", "verbatim", "versatile", "vindicate", "virtue", "vocabulary", "void", "voracious", "voyage", "vulnerable",
                                "wealth", "wedding", "wegde", "wheat", "whim", "wisdom", "wretch",
                                "xenophobe", "xylophone", "xylem",
                                "yarrow", "yeast", "yesterday", "yogurt", "yttrium", "yule", "youth", "yolk",
                                "zealot", "zealous", "zephyr", "zinc", "zombie", "zoology", "zymurgy", "zipper"}
    Dim poolLength As Integer = 0
    Dim rnd As Random = New Random
    Dim rand As Integer
    Dim letters() As Char
    Dim wordLength As Integer
    Dim counter As Integer = 0
    Dim maskedWord() As Char
    Dim letterHit As Integer = 0

    Public Function init()
        btn_m.Enabled = True
        btn_n.Enabled = True
        btn_b.Enabled = True
        btn_v.Enabled = True
        btn_c.Enabled = True
        btn_x.Enabled = True
        btn_z.Enabled = True
        btn_l.Enabled = True
        btn_k.Enabled = True
        btn_j.Enabled = True
        btn_h.Enabled = True
        btn_g.Enabled = True
        btn_f.Enabled = True
        btn_d.Enabled = True
        btn_s.Enabled = True
        btn_a.Enabled = True
        btn_p.Enabled = True
        btn_o.Enabled = True
        btn_i.Enabled = True
        btn_u.Enabled = True
        btn_y.Enabled = True
        btn_t.Enabled = True
        btn_r.Enabled = True
        btn_e.Enabled = True
        btn_w.Enabled = True
        btn_q.Enabled = True
        'set X's
        lbl_x1.Text = ""
        lbl_x2.Text = ""
        lbl_x3.Text = ""
        lbl_x4.Text = ""
        lbl_x5.Text = ""
        'counting words in the pool
        poolLength = 0
        For Each word As String In wordPool
            poolLength += 1
        Next
        'random selecting the index of the word from the pool to be solved
        Randomize()
        rand = rnd.Next(0, poolLength)
        'count word length
        wordLength = wordPool(rand).Length
        'store individual letters of words separately
        letters = wordPool(rand).ToUpper().ToCharArray()
        'masking the word
        ReDim maskedWord(wordLength - 1)
        counter = 0
        While counter < wordLength
            maskedWord(counter) = "_"
            counter += 1
        End While
        updateWord()
    End Function

    Public Function updateWord()
        'set masked word to display
        Select Case (wordLength)
            Case 11, 10
                lbl_ltr1.Text = maskedWord(0)
                lbl_ltr2.Text = maskedWord(1)
                lbl_ltr3.Text = maskedWord(2)
                lbl_ltr4.Text = maskedWord(3)
                lbl_ltr5.Text = maskedWord(4)
                lbl_ltr6.Text = maskedWord(5)
                lbl_ltr7.Text = maskedWord(6)
                lbl_ltr8.Text = maskedWord(7)
                lbl_ltr9.Text = maskedWord(8)
                lbl_ltr10.Text = maskedWord(9)
                If wordLength = 11 Then
                    lbl_ltr11.Text = maskedWord(10)
                Else
                    lbl_ltr11.Text = ""
                End If
            Case 9, 8
                lbl_ltr1.Text = ""
                lbl_ltr2.Text = maskedWord(0)
                lbl_ltr3.Text = maskedWord(1)
                lbl_ltr4.Text = maskedWord(2)
                lbl_ltr5.Text = maskedWord(3)
                lbl_ltr6.Text = maskedWord(4)
                lbl_ltr7.Text = maskedWord(5)
                lbl_ltr8.Text = maskedWord(6)
                lbl_ltr9.Text = maskedWord(7)
                If wordLength = 9 Then
                    lbl_ltr10.Text = maskedWord(8)
                Else
                    lbl_ltr10.Text = ""
                End If
                lbl_ltr11.Text = ""
            Case 7, 6
                lbl_ltr1.Text = ""
                lbl_ltr2.Text = ""
                lbl_ltr3.Text = maskedWord(0)
                lbl_ltr4.Text = maskedWord(1)
                lbl_ltr5.Text = maskedWord(2)
                lbl_ltr6.Text = maskedWord(3)
                lbl_ltr7.Text = maskedWord(4)
                lbl_ltr8.Text = maskedWord(5)
                If wordLength = 7 Then
                    lbl_ltr9.Text = maskedWord(6)
                Else
                    lbl_ltr9.Text = ""
                End If
                lbl_ltr10.Text = ""
                lbl_ltr11.Text = ""
            Case 5, 4
                lbl_ltr1.Text = ""
                lbl_ltr2.Text = ""
                lbl_ltr3.Text = ""
                lbl_ltr4.Text = maskedWord(0)
                lbl_ltr5.Text = maskedWord(1)
                lbl_ltr6.Text = maskedWord(2)
                lbl_ltr7.Text = maskedWord(3)
                If wordLength = 5 Then
                    lbl_ltr8.Text = maskedWord(4)
                Else
                    lbl_ltr8.Text = ""
                End If
                lbl_ltr9.Text = ""
                lbl_ltr10.Text = ""
                lbl_ltr11.Text = ""
            Case 3
                lbl_ltr1.Text = ""
                lbl_ltr2.Text = ""
                lbl_ltr3.Text = ""
                lbl_ltr4.Text = ""
                lbl_ltr5.Text = maskedWord(0)
                lbl_ltr6.Text = maskedWord(1)
                lbl_ltr7.Text = maskedWord(2)
                lbl_ltr8.Text = ""
                lbl_ltr9.Text = ""
                lbl_ltr10.Text = ""
                lbl_ltr11.Text = ""
            Case Else
                init()
        End Select
    End Function

    Public Function update(ByVal x)

        counter = 0
        letterHit = 0
        For Each letter As Char In letters
            If letter = x Then
                maskedWord(counter) = x
                updateWord()
                letterHit += 1
            End If
            counter += 1
        Next
        If letterHit > 0 Then
            'playSound(right)
        ElseIf lbl_x1.Text = "" Then
            lbl_x1.Text = "X"
            'playSound(err)
        ElseIf lbl_x2.Text = "" Then
            lbl_x2.Text = "X"
            'playSound(err)
        ElseIf lbl_x3.Text = "" Then
            lbl_x3.Text = "X"
            'playSound(err)
        ElseIf lbl_x4.Text = "" Then
            lbl_x4.Text = "X"
            'playSound(err)
        ElseIf lbl_x5.Text = "" Then
            lbl_x5.Text = "X"
            'playSound(err)
        End If
        status(check())
    End Function
    Public Function check() As String
        counter = 0
        For Each letter As Char In letters
            If lbl_x5.Text = "X" Then
                Return "Game Over"
            ElseIf maskedWord(counter) = "_" Then
                Return "Game On"
            Else
                counter += 1
            End If
        Next
        Return "Game Win"
    End Function

    Public Function status(ByVal x)
        Select Case x
            Case "Game Over"
                'playSound(Over)
                Dim response = MsgBox("You Lose! The word is " + wordPool(rand).ToUpper() & vbCrLf & " Play again?", MsgBoxStyle.YesNo, "Game Over!")
                If response = MsgBoxResult.Yes Then
                    init()
                Else
                    Me.Close()
                End If
            Case "Game Win"
                'playSound(Win)
                Dim response = MsgBox("You Won!" & vbCrLf & " Play again?", MsgBoxStyle.YesNo, "Congratulations!")
                If response = MsgBoxResult.Yes Then
                    init()
                Else
                    Me.Close()
                End If
        End Select
    End Function
    Private Sub frm_main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        init()
    End Sub


    Private Sub btn_q_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_q.Click
        btn_q.Enabled = False
        update("Q")
    End Sub

    Private Sub btn_w_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_w.Click
        btn_w.Enabled = False
        update("W")
    End Sub

    Private Sub btn_e_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_e.Click
        btn_e.Enabled = False
        update("E")
    End Sub

    Private Sub btn_r_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_r.Click
        btn_r.Enabled = False
        update("R")
    End Sub

    Private Sub btn_t_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_t.Click
        btn_t.Enabled = False
        update("T")
    End Sub

    Private Sub btn_y_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_y.Click
        btn_y.Enabled = False
        update("Y")
    End Sub

    Private Sub btn_u_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_u.Click
        btn_u.Enabled = False
        update("U")
    End Sub

    Private Sub btn_i_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_i.Click
        btn_i.Enabled = False
        update("I")
    End Sub

    Private Sub btn_o_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_o.Click
        btn_o.Enabled = False
        update("O")
    End Sub

    Private Sub btn_p_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_p.Click
        btn_p.Enabled = False
        update("P")
    End Sub

    Private Sub btn_a_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_a.Click
        btn_a.Enabled = False
        update("A")
    End Sub

    Private Sub btn_s_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_s.Click
        btn_s.Enabled = False
        update("S")
    End Sub

    Private Sub btn_d_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_d.Click
        btn_d.Enabled = False
        update("D")
    End Sub

    Private Sub btn_f_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_f.Click
        btn_f.Enabled = False
        update("F")
    End Sub

    Private Sub btn_g_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_g.Click
        btn_g.Enabled = False
        update("G")
    End Sub

    Private Sub btn_h_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_h.Click
        btn_h.Enabled = False
        update("H")
    End Sub

    Private Sub btn_j_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_j.Click
        btn_j.Enabled = False
        update("J")
    End Sub

    Private Sub btn_k_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_k.Click
        btn_k.Enabled = False
        update("K")
    End Sub

    Private Sub btn_l_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_l.Click
        btn_l.Enabled = False
        update("L")
    End Sub

    Private Sub btn_z_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_z.Click
        btn_z.Enabled = False
        update("Z")
    End Sub

    Private Sub btn_x_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_x.Click
        btn_x.Enabled = False
        update("X")
    End Sub

    Private Sub btn_c_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_c.Click
        btn_c.Enabled = False
        update("C")
    End Sub

    Private Sub btn_v_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_v.Click
        btn_v.Enabled = False
        update("V")
    End Sub

    Private Sub btn_b_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_b.Click
        btn_b.Enabled = False
        update("B")
    End Sub

    Private Sub btn_n_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_n.Click
        btn_n.Enabled = False
        update("N")
    End Sub

    Private Sub btn_m_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_m.Click
        btn_m.Enabled = False
        update("M")
    End Sub
End Class
