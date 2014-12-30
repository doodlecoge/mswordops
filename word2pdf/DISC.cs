using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Drawing;
using System.IO;

namespace word2pdf
{
	class DISC
	{

		public static string createDISCImage(
			int d0, int d1, int d2,
			int i0, int i1, int i2,
			int s0, int s1, int s2,
			int c0, int c1, int c2, 
			string discAxisImg)
		{

			#region axis

			Hashtable ht2 = new Hashtable();
			ht2["D24"] = new Rectangle(486, 82, 4, 4);
			ht2["D23"] = new Rectangle(486, 84, 4, 4);
			ht2["D22"] = new Rectangle(486, 86, 4, 4);
			ht2["D21"] = new Rectangle(486, 88, 4, 4);
			ht2["D20"] = new Rectangle(486, 90, 4, 4);
			ht2["D19"] = new Rectangle(486, 96, 4, 4);
			ht2["D18"] = new Rectangle(486, 102, 4, 4);
			ht2["D17"] = new Rectangle(486, 108, 4, 4);
			ht2["D16"] = new Rectangle(486, 110, 4, 4);
			ht2["D15"] = new Rectangle(486, 130, 4, 4);
			ht2["D14"] = new Rectangle(486, 150, 4, 4);
			ht2["D13"] = new Rectangle(486, 170, 4, 4);
			ht2["D12"] = new Rectangle(486, 190, 4, 4);
			ht2["D11"] = new Rectangle(486, 200, 4, 4);
			ht2["D10"] = new Rectangle(486, 210, 4, 4);
			ht2["D9"] = new Rectangle(486, 230, 4, 4);
			ht2["D8"] = new Rectangle(486, 250, 4, 4);
			ht2["D7"] = new Rectangle(486, 270, 4, 4);
			ht2["D6"] = new Rectangle(486, 290, 4, 4);
			ht2["D5"] = new Rectangle(486, 310, 4, 4);
			ht2["D4"] = new Rectangle(486, 320, 4, 4);
			ht2["D3"] = new Rectangle(486, 330, 4, 4);
			ht2["D2"] = new Rectangle(486, 340, 4, 4);
			ht2["D1"] = new Rectangle(486, 350, 4, 4);
			ht2["D0"] = new Rectangle(486, 370, 4, 4);
			ht2["D-1"] = new Rectangle(486, 390, 4, 4);
			ht2["D-2"] = new Rectangle(486, 410, 4, 4);
			ht2["D-3"] = new Rectangle(486, 420, 4, 4);
			ht2["D-4"] = new Rectangle(486, 430, 4, 4);
			ht2["D-5"] = new Rectangle(486, 440, 4, 4);
			ht2["D-6"] = new Rectangle(486, 450, 4, 4);
			ht2["D-7"] = new Rectangle(486, 460, 4, 4);
			ht2["D-8"] = new Rectangle(486, 460, 4, 4);
			ht2["D-9"] = new Rectangle(486, 470, 4, 4);
			ht2["D-10"] = new Rectangle(486, 490, 4, 4);
			ht2["D-11"] = new Rectangle(486, 510, 4, 4);
			ht2["D-12"] = new Rectangle(486, 520, 4, 4);
			ht2["D-13"] = new Rectangle(486, 530, 4, 4);
			ht2["D-14"] = new Rectangle(486, 550, 4, 4);
			ht2["D-15"] = new Rectangle(486, 560, 4, 4);
			ht2["D-16"] = new Rectangle(486, 560, 4, 4);
			ht2["D-17"] = new Rectangle(486, 560, 4, 4);
			ht2["D-18"] = new Rectangle(486, 560, 4, 4);
			ht2["D-19"] = new Rectangle(486, 560, 4, 4);
			ht2["D-20"] = new Rectangle(486, 560, 4, 4);
			ht2["D-21"] = new Rectangle(486, 570, 4, 4);
			ht2["D-22"] = new Rectangle(486, 570, 4, 4);
			ht2["D-23"] = new Rectangle(486, 570, 4, 4);
			ht2["D-24"] = new Rectangle(486, 570, 4, 4);


			ht2["I24"] = new Rectangle(544, 90, 4, 4);
			ht2["I23"] = new Rectangle(544, 90, 4, 4);
			ht2["I22"] = new Rectangle(544, 90, 4, 4);
			ht2["I21"] = new Rectangle(544, 90, 4, 4);
			ht2["I20"] = new Rectangle(544, 90, 4, 4);
			ht2["I19"] = new Rectangle(544, 90, 4, 4);
			ht2["I18"] = new Rectangle(544, 90, 4, 4);
			ht2["I17"] = new Rectangle(544, 90, 4, 4);
			ht2["I16"] = new Rectangle(544, 100, 4, 4);
			ht2["I15"] = new Rectangle(544, 100, 4, 4);
			ht2["I14"] = new Rectangle(544, 100, 4, 4);
			ht2["I13"] = new Rectangle(544, 100, 4, 4);
			ht2["I12"] = new Rectangle(544, 100, 4, 4);
			ht2["I11"] = new Rectangle(544, 100, 4, 4);
			ht2["I10"] = new Rectangle(544, 100, 4, 4);
			ht2["I9"] = new Rectangle(544, 110, 4, 4);
			ht2["I8"] = new Rectangle(544, 130, 4, 4);
			ht2["I7"] = new Rectangle(544, 150, 4, 4);
			ht2["I6"] = new Rectangle(544, 170, 4, 4);
			ht2["I5"] = new Rectangle(544, 210, 4, 4);
			ht2["I4"] = new Rectangle(544, 230, 4, 4);
			ht2["I3"] = new Rectangle(544, 270, 4, 4);
			ht2["I2"] = new Rectangle(544, 310, 4, 4);
			ht2["I1"] = new Rectangle(544, 330, 4, 4);
			ht2["I0"] = new Rectangle(544, 350, 4, 4);
			ht2["I-1"] = new Rectangle(544, 390, 4, 4);
			ht2["I-2"] = new Rectangle(544, 410, 4, 4);
			ht2["I-3"] = new Rectangle(544, 430, 4, 4);
			ht2["I-4"] = new Rectangle(544, 450, 4, 4);
			ht2["I-5"] = new Rectangle(544, 470, 4, 4);
			ht2["I-6"] = new Rectangle(544, 490, 4, 4);
			ht2["I-7"] = new Rectangle(544, 510, 4, 4);
			ht2["I-8"] = new Rectangle(544, 530, 4, 4);
			ht2["I-9"] = new Rectangle(544, 540, 4, 4);
			ht2["I-10"] = new Rectangle(544, 550, 4, 4);
			ht2["I-11"] = new Rectangle(544, 560, 4, 4);
			ht2["I-12"] = new Rectangle(544, 560, 4, 4);
			ht2["I-13"] = new Rectangle(544, 560, 4, 4);
			ht2["I-14"] = new Rectangle(544, 560, 4, 4);
			ht2["I-15"] = new Rectangle(544, 560, 4, 4);
			ht2["I-16"] = new Rectangle(544, 560, 4, 4);
			ht2["I-17"] = new Rectangle(544, 560, 4, 4);
			ht2["I-18"] = new Rectangle(544, 560, 4, 4);
			ht2["I-19"] = new Rectangle(544, 570, 4, 4);
			ht2["I-20"] = new Rectangle(544, 570, 4, 4);
			ht2["I-21"] = new Rectangle(544, 570, 4, 4);
			ht2["I-22"] = new Rectangle(544, 570, 4, 4);
			ht2["I-23"] = new Rectangle(544, 570, 4, 4);
			ht2["I-24"] = new Rectangle(544, 570, 4, 4);

			ht2["S24"] = new Rectangle(600, 80, 4, 4);
			ht2["S23"] = new Rectangle(600, 80, 4, 4);
			ht2["S22"] = new Rectangle(600, 80, 4, 4);
			ht2["S21"] = new Rectangle(600, 80, 4, 4);
			ht2["S20"] = new Rectangle(600, 80, 4, 4);
			ht2["S19"] = new Rectangle(600, 90, 4, 4);
			ht2["S18"] = new Rectangle(600, 100, 4, 4);
			ht2["S17"] = new Rectangle(600, 100, 4, 4);
			ht2["S16"] = new Rectangle(600, 100, 4, 4);
			ht2["S15"] = new Rectangle(600, 100, 4, 4);
			ht2["S14"] = new Rectangle(600, 100, 4, 4);
			ht2["S13"] = new Rectangle(600, 100, 4, 4);
			ht2["S12"] = new Rectangle(600, 100, 4, 4);
			ht2["S11"] = new Rectangle(600, 110, 4, 4);
			ht2["S10"] = new Rectangle(600, 130, 4, 4);
			ht2["S9"] = new Rectangle(600, 150, 4, 4);
			ht2["S8"] = new Rectangle(600, 170, 4, 4);
			ht2["S7"] = new Rectangle(600, 190, 4, 4);
			ht2["S6"] = new Rectangle(600, 210, 4, 4);
			ht2["S5"] = new Rectangle(600, 230, 4, 4);
			ht2["S4"] = new Rectangle(600, 250, 4, 4);
			ht2["S3"] = new Rectangle(600, 270, 4, 4);
			ht2["S2"] = new Rectangle(600, 290, 4, 4);
			ht2["S1"] = new Rectangle(600, 310, 4, 4);
			ht2["S0"] = new Rectangle(600, 330, 4, 4);
			ht2["S-1"] = new Rectangle(600, 370, 4, 4);
			ht2["S-2"] = new Rectangle(600, 410, 4, 4);
			ht2["S-3"] = new Rectangle(600, 420, 4, 4);
			ht2["S-4"] = new Rectangle(600, 430, 4, 4);
			ht2["S-5"] = new Rectangle(600, 450, 4, 4);
			ht2["S-6"] = new Rectangle(600, 460, 4, 4);
			ht2["S-7"] = new Rectangle(600, 470, 4, 4);
			ht2["S-8"] = new Rectangle(600, 490, 4, 4);
			ht2["S-9"] = new Rectangle(600, 510, 4, 4);
			ht2["S-10"] = new Rectangle(600, 530, 4, 4);
			ht2["S-11"] = new Rectangle(600, 540, 4, 4);
			ht2["S-12"] = new Rectangle(600, 550, 4, 4);
			ht2["S-13"] = new Rectangle(600, 560, 4, 4);
			ht2["S-14"] = new Rectangle(600, 560, 4, 4);
			ht2["S-15"] = new Rectangle(600, 560, 4, 4);
			ht2["S-16"] = new Rectangle(600, 560, 4, 4);
			ht2["S-17"] = new Rectangle(600, 560, 4, 4);
			ht2["S-18"] = new Rectangle(600, 560, 4, 4);
			ht2["S-19"] = new Rectangle(600, 570, 4, 4);
			ht2["S-20"] = new Rectangle(600, 580, 4, 4);
			ht2["S-21"] = new Rectangle(600, 580, 4, 4);
			ht2["S-22"] = new Rectangle(600, 580, 4, 4);
			ht2["S-23"] = new Rectangle(600, 580, 4, 4);
			ht2["S-24"] = new Rectangle(600, 580, 4, 4);


			ht2["C24"] = new Rectangle(660, 80, 4, 4);
			ht2["C23"] = new Rectangle(660, 80, 4, 4);
			ht2["C22"] = new Rectangle(660, 80, 4, 4);
			ht2["C21"] = new Rectangle(660, 80, 4, 4);
			ht2["C20"] = new Rectangle(660, 80, 4, 4);
			ht2["C19"] = new Rectangle(660, 80, 4, 4);
			ht2["C18"] = new Rectangle(660, 80, 4, 4);
			ht2["C17"] = new Rectangle(660, 80, 4, 4);
			ht2["C16"] = new Rectangle(660, 80, 4, 4);
			ht2["C15"] = new Rectangle(660, 90, 4, 4);
			ht2["C14"] = new Rectangle(660, 100, 4, 4);
			ht2["C13"] = new Rectangle(660, 100, 4, 4);
			ht2["C12"] = new Rectangle(660, 100, 4, 4);
			ht2["C11"] = new Rectangle(660, 100, 4, 4);
			ht2["C10"] = new Rectangle(660, 100, 4, 4);
			ht2["C9"] = new Rectangle(660, 100, 4, 4);
			ht2["C8"] = new Rectangle(660, 100, 4, 4);
			ht2["C7"] = new Rectangle(660, 110, 4, 4);
			ht2["C6"] = new Rectangle(660, 130, 4, 4);
			ht2["C5"] = new Rectangle(660, 150, 4, 4);
			ht2["C4"] = new Rectangle(660, 170, 4, 4);
			ht2["C3"] = new Rectangle(660, 210, 4, 4);
			ht2["C2"] = new Rectangle(660, 250, 4, 4);
			ht2["C1"] = new Rectangle(660, 270, 4, 4);
			ht2["C0"] = new Rectangle(660, 310, 4, 4);
			ht2["C-1"] = new Rectangle(660, 330, 4, 4);
			ht2["C-2"] = new Rectangle(660, 350, 4, 4);
			ht2["C-3"] = new Rectangle(660, 390, 4, 4);
			ht2["C-4"] = new Rectangle(660, 430, 4, 4);
			ht2["C-5"] = new Rectangle(660, 450, 4, 4);
			ht2["C-6"] = new Rectangle(660, 460, 4, 4);
			ht2["C-7"] = new Rectangle(660, 470, 4, 4);
			ht2["C-8"] = new Rectangle(660, 490, 4, 4);
			ht2["C-9"] = new Rectangle(660, 510, 4, 4);
			ht2["C-10"] = new Rectangle(660, 530, 4, 4);
			ht2["C-11"] = new Rectangle(660, 540, 4, 4);
			ht2["C-12"] = new Rectangle(660, 550, 4, 4);
			ht2["C-13"] = new Rectangle(660, 560, 4, 4);
			ht2["C-14"] = new Rectangle(660, 560, 4, 4);
			ht2["C-15"] = new Rectangle(660, 560, 4, 4);
			ht2["C-16"] = new Rectangle(660, 570, 4, 4);
			ht2["C-17"] = new Rectangle(660, 570, 4, 4);
			ht2["C-18"] = new Rectangle(660, 570, 4, 4);
			ht2["C-19"] = new Rectangle(660, 570, 4, 4);
			ht2["C-20"] = new Rectangle(660, 570, 4, 4);
			ht2["C-21"] = new Rectangle(660, 570, 4, 4);
			ht2["C-22"] = new Rectangle(660, 570, 4, 4);
			ht2["C-23"] = new Rectangle(660, 570, 4, 4);
			ht2["C-24"] = new Rectangle(660, 570, 4, 4);

			Hashtable ht0 = new Hashtable();
			ht0["D0"] = new Rectangle(30, 570, 4, 4);
			ht0["D1"] = new Rectangle(30, 530, 4, 4);
			ht0["D2"] = new Rectangle(30, 510, 4, 4);
			ht0["D3"] = new Rectangle(30, 470, 4, 4);
			ht0["D4"] = new Rectangle(30, 430, 4, 4);
			ht0["D5"] = new Rectangle(30, 410, 4, 4);
			ht0["D6"] = new Rectangle(30, 370, 4, 4);
			ht0["D7"] = new Rectangle(30, 350, 4, 4);
			ht0["D8"] = new Rectangle(30, 310, 4, 4);
			ht0["D9"] = new Rectangle(30, 290, 4, 4);
			ht0["D10"] = new Rectangle(30, 270, 4, 4);
			ht0["D11"] = new Rectangle(30, 250, 4, 4);
			ht0["D12"] = new Rectangle(30, 230, 4, 4);
			ht0["D13"] = new Rectangle(30, 210, 4, 4);
			ht0["D14"] = new Rectangle(30, 190, 4, 4);
			ht0["D15"] = new Rectangle(30, 110, 4, 4);
			ht0["D16"] = new Rectangle(30, 90, 4, 4);
			ht0["D17"] = new Rectangle(30, 80, 4, 4);
			ht0["D18"] = new Rectangle(30, 80, 4, 4);
			ht0["D19"] = new Rectangle(30, 80, 4, 4);
			ht0["D20"] = new Rectangle(30, 70, 4, 4);
			ht0["D21"] = new Rectangle(30, 60, 4, 4);
			ht0["D22"] = new Rectangle(30, 60, 4, 4);
			ht0["D23"] = new Rectangle(30, 60, 4, 4);
			ht0["D24"] = new Rectangle(30, 60, 4, 4);
			ht0["I0"] = new Rectangle(88, 550, 4, 4);
			ht0["I1"] = new Rectangle(88, 510, 4, 4);
			ht0["I2"] = new Rectangle(88, 430, 4, 4);
			ht0["I3"] = new Rectangle(88, 390, 4, 4);
			ht0["I4"] = new Rectangle(88, 330, 4, 4);
			ht0["I5"] = new Rectangle(88, 290, 4, 4);
			ht0["I6"] = new Rectangle(88, 250, 4, 4);
			ht0["I7"] = new Rectangle(88, 210, 4, 4);
			ht0["I8"] = new Rectangle(88, 190, 4, 4);
			ht0["I9"] = new Rectangle(88, 130, 4, 4);
			ht0["I10"] = new Rectangle(88, 90, 4, 4);
			ht0["I11"] = new Rectangle(88, 80, 4, 4);
			ht0["I12"] = new Rectangle(88, 80, 4, 4);
			ht0["I13"] = new Rectangle(88, 80, 4, 4);
			ht0["I14"] = new Rectangle(88, 80, 4, 4);
			ht0["I15"] = new Rectangle(88, 80, 4, 4);
			ht0["I16"] = new Rectangle(88, 80, 4, 4);
			ht0["I17"] = new Rectangle(88, 70, 4, 4);
			ht0["I18"] = new Rectangle(88, 60, 4, 4);
			ht0["I19"] = new Rectangle(88, 60, 4, 4);
			ht0["I20"] = new Rectangle(88, 60, 4, 4);
			ht0["I21"] = new Rectangle(88, 60, 4, 4);
			ht0["I22"] = new Rectangle(88, 60, 4, 4);
			ht0["I23"] = new Rectangle(88, 60, 4, 4);
			ht0["I24"] = new Rectangle(88, 60, 4, 4);
			ht0["S0"] = new Rectangle(144, 530, 4, 4);
			ht0["S1"] = new Rectangle(144, 490, 4, 4);
			ht0["S2"] = new Rectangle(144, 470, 4, 4);
			ht0["S3"] = new Rectangle(144, 430, 4, 4);
			ht0["S4"] = new Rectangle(144, 390, 4, 4);
			ht0["S5"] = new Rectangle(144, 350, 4, 4);
			ht0["S6"] = new Rectangle(144, 330, 4, 4);
			ht0["S7"] = new Rectangle(144, 290, 4, 4);
			ht0["S8"] = new Rectangle(144, 270, 4, 4);
			ht0["S9"] = new Rectangle(144, 230, 4, 4);
			ht0["S10"] = new Rectangle(144, 190, 4, 4);
			ht0["S11"] = new Rectangle(144, 150, 4, 4);
			ht0["S12"] = new Rectangle(144, 90, 4, 4);
			ht0["S13"] = new Rectangle(144, 80, 4, 4);
			ht0["S14"] = new Rectangle(144, 80, 4, 4);
			ht0["S15"] = new Rectangle(144, 80, 4, 4);
			ht0["S16"] = new Rectangle(144, 80, 4, 4);
			ht0["S17"] = new Rectangle(144, 80, 4, 4);
			ht0["S18"] = new Rectangle(144, 80, 4, 4);
			ht0["S19"] = new Rectangle(144, 70, 4, 4);
			ht0["S20"] = new Rectangle(144, 60, 4, 4);
			ht0["S21"] = new Rectangle(144, 60, 4, 4);
			ht0["S22"] = new Rectangle(144, 60, 4, 4);
			ht0["S23"] = new Rectangle(144, 60, 4, 4);
			ht0["S24"] = new Rectangle(144, 60, 4, 4);
			ht0["C0"] = new Rectangle(202, 570, 4, 4);
			ht0["C1"] = new Rectangle(202, 520, 4, 4);
			ht0["C2"] = new Rectangle(202, 470, 4, 4);
			ht0["C3"] = new Rectangle(202, 410, 4, 4);
			ht0["C4"] = new Rectangle(202, 350, 4, 4);
			ht0["C5"] = new Rectangle(202, 290, 4, 4);
			ht0["C6"] = new Rectangle(202, 250, 4, 4);
			ht0["C7"] = new Rectangle(202, 210, 4, 4);
			ht0["C8"] = new Rectangle(202, 170, 4, 4);
			ht0["C9"] = new Rectangle(202, 90, 4, 4);
			ht0["C10"] = new Rectangle(202, 80, 4, 4);
			ht0["C11"] = new Rectangle(202, 80, 4, 4);
			ht0["C12"] = new Rectangle(202, 80, 4, 4);
			ht0["C13"] = new Rectangle(202, 80, 4, 4);
			ht0["C14"] = new Rectangle(202, 80, 4, 4);
			ht0["C15"] = new Rectangle(202, 70, 4, 4);
			ht0["C16"] = new Rectangle(202, 60, 4, 4);
			ht0["C17"] = new Rectangle(202, 60, 4, 4);
			ht0["C18"] = new Rectangle(202, 60, 4, 4);
			ht0["C19"] = new Rectangle(202, 60, 4, 4);
			ht0["C20"] = new Rectangle(202, 60, 4, 4);
			ht0["C21"] = new Rectangle(202, 60, 4, 4);
			ht0["C22"] = new Rectangle(202, 60, 4, 4);
			ht0["C23"] = new Rectangle(202, 60, 4, 4);
			ht0["C24"] = new Rectangle(202, 60, 4, 4);

			Hashtable ht1 = new Hashtable();
			ht1["D0"] = new Rectangle(258, 70, 4, 4);
			ht1["D1"] = new Rectangle(258, 130, 4, 4);
			ht1["D2"] = new Rectangle(258, 230, 4, 4);
			ht1["D3"] = new Rectangle(258, 290, 4, 4);
			ht1["D4"] = new Rectangle(258, 330, 4, 4);
			ht1["D5"] = new Rectangle(258, 350, 4, 4);
			ht1["D6"] = new Rectangle(258, 370, 4, 4);
			ht1["D7"] = new Rectangle(258, 390, 4, 4);
			ht1["D8"] = new Rectangle(258, 410, 4, 4);
			ht1["D9"] = new Rectangle(258, 430, 4, 4);
			ht1["D10"] = new Rectangle(258, 450, 4, 4);
			ht1["D11"] = new Rectangle(258, 470, 4, 4);
			ht1["D12"] = new Rectangle(258, 490, 4, 4);
			ht1["D13"] = new Rectangle(258, 510, 4, 4);
			ht1["D14"] = new Rectangle(258, 520, 4, 4);
			ht1["D15"] = new Rectangle(258, 530, 4, 4);
			ht1["D16"] = new Rectangle(258, 550, 4, 4);
			ht1["D17"] = new Rectangle(258, 560, 4, 4);
			ht1["D18"] = new Rectangle(258, 560, 4, 4);
			ht1["D19"] = new Rectangle(258, 560, 4, 4);
			ht1["D20"] = new Rectangle(258, 560, 4, 4);
			ht1["D21"] = new Rectangle(258, 570, 4, 4);
			ht1["D22"] = new Rectangle(258, 580, 4, 4);
			ht1["D23"] = new Rectangle(258, 580, 4, 4);
			ht1["D24"] = new Rectangle(258, 580, 4, 4);

			ht1["I0"] = new Rectangle(316, 70, 4, 4);
			ht1["I1"] = new Rectangle(316, 130, 4, 4);
			ht1["I2"] = new Rectangle(316, 230, 4, 4);
			ht1["I3"] = new Rectangle(316, 290, 4, 4);
			ht1["I4"] = new Rectangle(316, 330, 4, 4);
			ht1["I5"] = new Rectangle(316, 370, 4, 4);
			ht1["I6"] = new Rectangle(316, 410, 4, 4);
			ht1["I7"] = new Rectangle(316, 450, 4, 4);
			ht1["I8"] = new Rectangle(316, 490, 4, 4);
			ht1["I9"] = new Rectangle(316, 510, 4, 4);
			ht1["I10"] = new Rectangle(316, 520, 4, 4);
			ht1["I11"] = new Rectangle(316, 530, 4, 4);
			ht1["I12"] = new Rectangle(316, 550, 4, 4);
			ht1["I13"] = new Rectangle(316, 550, 4, 4);
			ht1["I14"] = new Rectangle(316, 550, 4, 4);
			ht1["I15"] = new Rectangle(316, 550, 4, 4);
			ht1["I16"] = new Rectangle(316, 550, 4, 4);
			ht1["I17"] = new Rectangle(316, 550, 4, 4);
			ht1["I18"] = new Rectangle(316, 550, 4, 4);
			ht1["I19"] = new Rectangle(316, 570, 4, 4);
			ht1["I20"] = new Rectangle(316, 580, 4, 4);
			ht1["I21"] = new Rectangle(316, 580, 4, 4);
			ht1["I22"] = new Rectangle(316, 580, 4, 4);
			ht1["I23"] = new Rectangle(316, 580, 4, 4);
			ht1["I24"] = new Rectangle(316, 580, 4, 4);
			ht1["S0"] = new Rectangle(372, 70, 4, 4);
			ht1["S1"] = new Rectangle(372, 90, 4, 4);
			ht1["S2"] = new Rectangle(372, 130, 4, 4);
			ht1["S3"] = new Rectangle(372, 250, 4, 4);
			ht1["S4"] = new Rectangle(372, 290, 4, 4);
			ht1["S5"] = new Rectangle(372, 330, 4, 4);
			ht1["S6"] = new Rectangle(372, 350, 4, 4);
			ht1["S7"] = new Rectangle(372, 390, 4, 4);
			ht1["S8"] = new Rectangle(372, 410, 4, 4);
			ht1["S9"] = new Rectangle(372, 450, 4, 4);
			ht1["S10"] = new Rectangle(372, 490, 4, 4);
			ht1["S11"] = new Rectangle(372, 510, 4, 4);
			ht1["S12"] = new Rectangle(372, 530, 4, 4);
			ht1["S13"] = new Rectangle(372, 550, 4, 4);
			ht1["S14"] = new Rectangle(372, 560, 4, 4);
			ht1["S15"] = new Rectangle(372, 560, 4, 4);
			ht1["S16"] = new Rectangle(372, 560, 4, 4);
			ht1["S17"] = new Rectangle(372, 560, 4, 4);
			ht1["S18"] = new Rectangle(372, 560, 4, 4);
			ht1["S19"] = new Rectangle(372, 570, 4, 4);
			ht1["S20"] = new Rectangle(372, 580, 4, 4);
			ht1["S21"] = new Rectangle(372, 580, 4, 4);
			ht1["S22"] = new Rectangle(372, 580, 4, 4);
			ht1["S23"] = new Rectangle(372, 580, 4, 4);
			ht1["S24"] = new Rectangle(372, 580, 4, 4);
			ht1["C0"] = new Rectangle(430, 70, 4, 4);
			ht1["C1"] = new Rectangle(430, 90, 4, 4);
			ht1["C2"] = new Rectangle(430, 170, 4, 4);
			ht1["C3"] = new Rectangle(430, 270, 4, 4);
			ht1["C4"] = new Rectangle(430, 310, 4, 4);
			ht1["C5"] = new Rectangle(430, 330, 4, 4);
			ht1["C6"] = new Rectangle(430, 350, 4, 4);
			ht1["C7"] = new Rectangle(430, 370, 4, 4);
			ht1["C8"] = new Rectangle(430, 410, 4, 4);
			ht1["C9"] = new Rectangle(430, 450, 4, 4);
			ht1["C10"] = new Rectangle(430, 490, 4, 4);
			ht1["C11"] = new Rectangle(430, 510, 4, 4);
			ht1["C12"] = new Rectangle(430, 530, 4, 4);
			ht1["C13"] = new Rectangle(430, 550, 4, 4);
			ht1["C14"] = new Rectangle(430, 560, 4, 4);
			ht1["C15"] = new Rectangle(430, 560, 4, 4);
			ht1["C16"] = new Rectangle(430, 570, 4, 4);
			ht1["C17"] = new Rectangle(430, 580, 4, 4);
			ht1["C18"] = new Rectangle(430, 580, 4, 4);
			ht1["C19"] = new Rectangle(430, 580, 4, 4);
			ht1["C20"] = new Rectangle(430, 580, 4, 4);
			ht1["C21"] = new Rectangle(430, 580, 4, 4);
			ht1["C22"] = new Rectangle(430, 580, 4, 4);
			ht1["C23"] = new Rectangle(430, 580, 4, 4);
			ht1["C24"] = new Rectangle(430, 580, 4, 4);
			#endregion

			try
			{
				//string path = Path.GetFullPath("DISC.jpg");
				string path = Path.GetFullPath(discAxisImg);
				Bitmap source = new Bitmap(path);
				Graphics target = Graphics.FromImage(source);
				Brush brush = new SolidBrush(Color.Red);
				Pen pen = new Pen(brush, 4);
				pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;

				target.DrawEllipse(pen, (Rectangle)ht0["D" + d0]);
				target.DrawEllipse(pen, (Rectangle)ht0["I" + i0]);
				target.DrawEllipse(pen, (Rectangle)ht0["S" + s0]);
				target.DrawEllipse(pen, (Rectangle)ht0["C" + c0]);
				target.DrawLine(pen,
					((Rectangle)ht0["D" + d0]).X,
					((Rectangle)ht0["D" + d0]).Y,
					((Rectangle)ht0["I" + i0]).X,
					((Rectangle)ht0["I" + i0]).Y
					);
				target.DrawLine(pen,
					((Rectangle)ht0["I" + i0]).X,
					((Rectangle)ht0["I" + i0]).Y,
					((Rectangle)ht0["S" + s0]).X,
					((Rectangle)ht0["S" + s0]).Y
					);
				target.DrawLine(pen,
					((Rectangle)ht0["S" + s0]).X,
					((Rectangle)ht0["S" + s0]).Y,
					((Rectangle)ht0["C" + c0]).X,
					((Rectangle)ht0["C" + c0]).Y
					);

				target.DrawEllipse(pen, (Rectangle)ht1["D" + d1]);
				target.DrawEllipse(pen, (Rectangle)ht1["I" + i1]);
				target.DrawEllipse(pen, (Rectangle)ht1["S" + s1]);
				target.DrawEllipse(pen, (Rectangle)ht1["C" + c1]);
				target.DrawLine(pen,
					((Rectangle)ht1["D" + d1]).X,
					((Rectangle)ht1["D" + d1]).Y,
					((Rectangle)ht1["I" + i1]).X,
					((Rectangle)ht1["I" + i1]).Y
					);
				target.DrawLine(pen,
					((Rectangle)ht1["I" + i1]).X,
					((Rectangle)ht1["I" + i1]).Y,
					((Rectangle)ht1["S" + s1]).X,
					((Rectangle)ht1["S" + s1]).Y
					);
				target.DrawLine(pen,
					((Rectangle)ht1["S" + s1]).X,
					((Rectangle)ht1["S" + s1]).Y,
					((Rectangle)ht1["C" + c1]).X,
					((Rectangle)ht1["C" + c1]).Y
					);

				target.DrawEllipse(pen, (Rectangle)ht2["D" + d2]);
				target.DrawEllipse(pen, (Rectangle)ht2["I" + i2]);
				target.DrawEllipse(pen, (Rectangle)ht2["S" + s2]);
				target.DrawEllipse(pen, (Rectangle)ht2["C" + c2]);
				target.DrawLine(pen,
					((Rectangle)ht2["D" + d2]).X,
					((Rectangle)ht2["D" + d2]).Y,
					((Rectangle)ht2["I" + i2]).X,
					((Rectangle)ht2["I" + i2]).Y
					);
				target.DrawLine(pen,
					((Rectangle)ht2["I" + i2]).X,
					((Rectangle)ht2["I" + i2]).Y,
					((Rectangle)ht2["S" + s2]).X,
					((Rectangle)ht2["S" + s2]).Y
					);
				target.DrawLine(pen,
					((Rectangle)ht2["S" + s2]).X,
					((Rectangle)ht2["S" + s2]).Y,
					((Rectangle)ht2["C" + c2]).X,
					((Rectangle)ht2["C" + c2]).Y
					);


				string img = Path.GetFullPath(DateTime.Now.Ticks + ".jpg");
				source.Save(img, System.Drawing.Imaging.ImageFormat.Jpeg);
				return img;
			}
			catch (Exception)
			{
				//Console.Out.WriteLine(ex.Message);
				//Console.Out.WriteLine(ex.StackTrace);
				return null;
			}
		}

	}
}
