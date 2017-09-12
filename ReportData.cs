namespace TableHandlers
{
    
    public static class ReportData
    {
        public static object[][] LoadDataDemo1()
        {
            //string floatToStrAccuracy = "N5";
            string[][] Result = new string[0][];
            var units = DemoEntities.CreateSC();
            /*
            table.addColumn("Unit").EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues().UseRowDataAsInfoBlockDelimeter();
             table.addColumn("Health Points").EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
             table.addColumn("Shield Points").EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
             table.addColumn("Description").EnableAutoMergeSameValues();
             table.addColumn("Damage [splash radius]").EnableAutoMergeSameValues();
             table.addColumn("Range").EnableAutoMergeSameValues();
             table.addColumn("Cooldown").EnableAutoMergeSameValues();
             table.addColumn("Minerals").EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues(); 
             table.addColumn("Vespene").EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues(); 
             table.addColumn("Resources").EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues(); 
             table.addColumn("#num");
             +Weapon type?
             +Has regen?
             +Splash
             */
            var demoString = "SAME_VALUE";

            var offset = 0;
            int total_counter = 1;
            const int total_fields_count = 15;
            object[][] dataRows = new object[
                    0
                    ][];

            foreach (var unit in units)
            {
                if (unit.weapons != null)
                {
                    for (int i = 0; i < unit.weapons.Length; i++)
                    {
                        object[] row = new object[total_fields_count];
                        for (int j = 0; j < row.Length; j++) { row[j] = null; }
                        row[3] = unit.weapons[i].name;
                        row[4] = unit.weapons[i].damage;
                        row[5] = unit.weapons[i].range;
                        row[6] = unit.weapons[i].cooldown;
                        row[12] = unit.weapons[i].type;
                        row[13] = unit.weapons[i].splashRadius;
                        row[14] = demoString;

                        Utilities.add<object[]>(ref dataRows, row);
                    }
                    offset += unit.weapons.Length;
                }
                else
                {
                    Utilities.add<object[]>(ref dataRows, new object[total_fields_count]);
                    offset++;
                }

                var index = (unit.weapons != null) ? offset - unit.weapons.Length : offset-1;
                dataRows[index][0] = unit.name;
                dataRows[index][1] = unit.healthPoints;
                dataRows[index][2] = (unit.hasShields) ? (object)unit.shieldPoints : null;
                dataRows[index][7] = unit.mineralsCost;
                dataRows[index][8] = unit.vespeneCost;
                dataRows[index][9] = unit.resourceCost;
                dataRows[index][10] = total_counter;
                dataRows[index][11] = unit.hasRegen;
                dataRows[index][14] = demoString;

                total_counter++;
            }

            return dataRows;
        }



        public static object[][] LoadDataDemo2()
        {
            var Rw1 = new string[] { "red", "red", "1", "" };
            var Rw2 = new string[] { "red", "9", "2", "" };
            var Rw3 = new string[] { "0", "1", "2", "---" };
            var Rw4 = new string[] { "0", "1", "2", "" };

            return (new string[4][] { Rw1, Rw2, Rw3, Rw4 }) as object[][];

        }




        public static object[][] LoadDataDemo3()
        {
            var Rw1 = new string[] { "Al", "1.5", null, "OOP" };
            var Rw2 = new string[] { "El", "3.4", null, "OPO" };
            var Rw3 = new string[] { "KI", "7", null, "PPO" };

            return (new string[3][] { Rw1, Rw2, Rw3 }) as object[][];

        }

    }





    public class DemoEntities
    {
        public class Unit
        {
            public string name;

            public int healthPoints;
            public bool hasRegen;
            public bool hasShields;
            public int shieldPoints;

            public Weaponry[] weapons;


            public int mineralsCost;
            public int vespeneCost;
            public int resourceCost;

            public Unit(string name, int healthPoints, bool hasRegen, bool hasShields, int shieldPoints, int mineralsCost, int vespeneCost, int resourceCost)
            {
                this.name = name;
                this.healthPoints = healthPoints;
                this.hasRegen = hasRegen;
                this.hasShields = hasShields;
                this.shieldPoints = shieldPoints;
                this.mineralsCost = mineralsCost;
                this.vespeneCost = vespeneCost;
                this.resourceCost = resourceCost;
            }

            public Unit addWeapon(Weaponry wpn)
            {
                Utilities.add<Weaponry>(ref weapons, wpn);
                return this;
            }
        }

        public class Weaponry
        {
            public static class Type
            {
                public const int AA = 1;
                public const int AG = 2;
                public const int AU = 3;
            }

            public int type;
            public string name;
            public int damage;
            public bool hasSplash;
            public float splashRadius;
            public float range;
            public float cooldown;

            public Weaponry(int type, string name, int damage, bool hasSplash, float splashRadius, float range, float cooldown)
            {
                this.type = type;
                this.name = name;
                this.damage = damage;
                this.hasSplash = hasSplash;
                this.splashRadius = splashRadius;
                this.range = range;
                this.cooldown = cooldown;
            }
        }

        public static Unit[] CreateSC()
        {
            Unit[] SC_Units = new Unit[0];

            Utilities.add<Unit>(ref SC_Units,
                    new Unit("Wraith", 120, false, false, 0, 150, 100, 2)
                        .addWeapon(new Weaponry(Weaponry.Type.AA, "Homin missiles", 20, false, 0, (float)7.66, (float)0.8))
                        .addWeapon(new Weaponry(Weaponry.Type.AG, "Burst Lasers", 8, false, 0, (float)7.01, 1))
                );


            Utilities.add<Unit>(ref SC_Units,
                    new Unit("Battlecruiser", 400, false, false, 0, 400, 300, 4)
                        .addWeapon(new Weaponry(Weaponry.Type.AA, "ATS Lasers", 25, false, 0, (float)10.34, (float)1.34))
                        .addWeapon(new Weaponry(Weaponry.Type.AG, "ATA Lasers", 25, false, 0, (float)10.35, (float)1.34))
                );

            Utilities.add<Unit>(ref SC_Units,
                    new Unit("Science vessel", 200, false, false, 0, 100, 225, 2)
                );

            Utilities.add<Unit>(ref SC_Units,
                    new Unit("Scout", 150, false, true, 100, 200, 100, 3)
                        .addWeapon(new Weaponry(Weaponry.Type.AA, "Blaster torpedoes", 24, false, 0, (float)8.21, 1))
                        .addWeapon(new Weaponry(Weaponry.Type.AG, "Photon blasters", 8, false, 0, (float)7.92, (float)1.1))
                );

            Utilities.add<Unit>(ref SC_Units,
                    new Unit("Corsair", 80, false, true, 100, 150, 100, 2)
                        .addWeapon(new Weaponry(Weaponry.Type.AA, "Warp lasers", 8, true, (float)1.23, (float)5.66, (float)0.5))
                );

            Utilities.add<Unit>(ref SC_Units,
                    new Unit("Mutalisk", 120, true, false, 0, 100, 100, 2)
                        .addWeapon(new Weaponry(Weaponry.Type.AA, "Flying blade", 9, false, 0, (float)6.18, (float)0.84))
                        .addWeapon(new Weaponry(Weaponry.Type.AG, "Flying blade", 9, false, 0, (float)6.28, (float)0.84))
                );

            Utilities.add<Unit>(ref SC_Units,
                    new Unit("Overlord", 100, true, false, 0, 100, 0, -8)
                );


            return SC_Units;

        }

    }
}
