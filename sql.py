import pymysql


# from REG_to_APPLY import Registration_sadik


class Database:
    def __init__(self):
        self.conn = pymysql.connect(
            user='root',
            password='ntygazRPNautoz',
            host='127.0.0.1',
            port=3307,
            database='knm'
        )

    def get_terr_upravlenie_name(self, condition):
        with self.conn.cursor() as cursor:
            cursor.execute(f"SELECT name FROM knd_terr_upravlenie WHERE name='{condition}';")
            result = cursor.fetchall()
            return result

    def get_terr_upravlenie_id(self, condition):
        with self.conn.cursor() as cursor:
            cursor.execute(f"SELECT id FROM knd_terr_upravlenie WHERE name='{condition}';")
            result = cursor.fetchall()
            if result != ():
                return result[0][0]
            else:
                return result

    def get_plan_proverok_id(self, condition):
        with self.conn.cursor() as cursor:
            cursor.execute(f"SELECT id FROM knd_plan_proverok WHERE number='{condition}';")
            result = cursor.fetchall()
            if result != ():
                return result[0][0]
            else:
                return result

    def get_kind_inspection_id(self, condition):
        with self.conn.cursor() as cursor:
            cursor.execute(f"SELECT id FROM knd_kind_inspection WHERE kind='{condition}';")
            result = cursor.fetchall()
            if result != ():
                return result[0][0]
            else:
                return result

    def get_inspection_id(self, condition):
        with self.conn.cursor() as cursor:
            cursor.execute(f"SELECT id FROM knd_inspection WHERE number='{condition}';")
            result = cursor.fetchall()
            if result != ():
                return result[0][0]
            else:
                return result

    def insert_plan_proverok(self, ter_upr, number, year, count, status):
        with self.conn.cursor() as cursor:
            cursor.execute(
                f"""INSERT INTO knd_plan_proverok(ter_upr_id, number, year, count, status) VALUES ('{ter_upr}', '{number}', '{year}', '{count}', '{status}')""")
            self.conn.commit()

    def insert_terr_uprav(self, name):
        with self.conn.cursor() as cursor:
            cursor.execute(f"""INSERT INTO knd_terr_upravlenie(name) VALUES ('{name}')""")
            self.conn.commit()

    def insert_subject_with_return_id(self, name, address, inn, ogrn):
        with self.conn.cursor() as cursor:
            cursor.execute(
                f"""INSERT INTO knd_subject(name, address, e_mail, district, inn, ogrn) VALUES ('{name}', '{address}', '""', '""', '{inn}', '{ogrn}')""")
            self.conn.commit()
            cursor.execute(f"""SELECT LAST_INSERT_ID();""")
            result = cursor.fetchall()
            if result != ():
                return result[0][0]
            else:
                return result

    def insert_object_with_return_id(self, subject_id, kind, address, risk_id):
        with self.conn.cursor() as cursor:
            cursor.execute(
                f"""INSERT INTO knd_object(subject_id, kind, address, risk_id) VALUES ('{subject_id}', '{kind}', '{address}', '{risk_id}')""")
            self.conn.commit()

            cursor.execute(f"""SELECT LAST_INSERT_ID();""")
            result = cursor.fetchall()
            if result != ():
                return result[0][0]
            else:
                return result

    def insert_inspection_with_return_id(self, plan_id: int, kind_id: int, number):
        with self.conn.cursor() as cursor:
            cursor.execute(
                f"""INSERT INTO knd_inspection(plan_id, kind_id, number) VALUES ('{plan_id}', '{kind_id}', '{number}')""")
            self.conn.commit()

            cursor.execute(f"""SELECT LAST_INSERT_ID();""")
            result = cursor.fetchall()
            if result != ():
                return result[0][0]
            else:
                return result

    def insert_m_to_m_object_inspection(self, inspection_id, object_id):
        with self.conn.cursor() as cursor:
            cursor.execute(
                f"""INSERT INTO knd_m_to_m_object_inspection(inspection_id, object_id) VALUES ('{inspection_id}', '{object_id}')""")
            self.conn.commit()


    def find_risk_id(self, condition):
        with self.conn.cursor() as cursor:
            cursor.execute(f"SELECT id FROM knd_risk_category WHERE category='{condition}';")
            result = cursor.fetchall()
            if result != ():
                return result[0][0]
            else:
                return result

    def change_names_knd_terr_upravlenie(self, terr_id, new_name):
        with self.conn.cursor() as cursor:
            cursor.execute(f"UPDATE knd_terr_upravlenie SET name='{new_name}' where id={terr_id}")
            self.conn.commit()

    def get_connect(self):

        rang = self.get_info()

        # print(rang[0][12])
        # self.listrang(rang)
        self.raspred(rang)
        # self.get_sadik_info(31)

    def raspred(self, rang):
        ready = []
        applied = []
        registred = []
        list = {'ready': {'conteiner': ready, 'function': 'будем согласовывать'},
                'applied': {'conteiner': applied, 'function': 'будем регистрировать'},
                'registred': {'conteiner': registred, 'function': 'будем отправлять'}}
        for o in rang:
            status = o[11]
            for k, v in list.items():
                if status == k:
                    print(f'{v["function"]}  {o[3]}')
                    v["conteiner"].append(o)
        print(f'ready {ready}')
        print(f'applied {applied}')
        print(f'registred {registred}')
        for i in ready:
            print(i)

    def listrang(self, rang):
        for o in rang:
            id_o = o[0]
            groups = o[1]
            group_size = o[2]
            date_start = o[3]
            date_end = o[4]
            reason = o[5]
            fio_covid = o[6]
            fio_post = o[7]
            last_day = o[8]
            identify_day = o[9]
            address_spe = o[10]
            status = o[11]
            sadik_id = o[14]

            sadik = self.get_sadik_info(sadik_id)
            # print(sadik[0])
            id_s = sadik[0][0]
            district = sadik[0][1]
            properties = sadik[0][2]
            only_name = sadik[0][3]
            address = sadik[0][4]
            fio_director = sadik[0][5]
            e_mail = sadik[0][6]
            inn = sadik[0][7]
            ogrn = sadik[0][8]
            sed_name = sadik[0][9]

            print(o)

            # self.level_up_status(id_o, status)

    def get_sadik_info(self, sadik_id):
        with self.conn.cursor() as cursor:
            cursor.execute(f"Select * FROM sadidi_sadik where id='{sadik_id}'")
            result = cursor.fetchall()
            return result

    # def get_info(self):
    #     # self.cursor.execute("Select * FROM sadidi_ordinary where status='ready'")
    #     with self.conn.cursor() as cursor:
    #         cursor.execute("Select * FROM sadidi_ordinary")
    #         result = cursor.fetchall()
    #         return result

    def level_up_status(self, id, status):
        referense = {'ready': 'applied', 'applied': 'registred', 'registred': 'sended', 'sended': 'ready'}
        # print(referense)
        if status == 'sended':
            return
        with self.conn.cursor() as cursor:
            for r in referense:
                if status == r:
                    # print(f'Новый статус:  {referense[r]}')
                    cursor.execute(f"UPDATE sadidi_ordinary SET status='{referense[r]}' where id={id}")
                    self.conn.commit()
                    return

    def assign_number(self, id, number):
        with self.conn.cursor() as cursor:
            cursor.execute(f"UPDATE sadidi_ordinary SET doc_number='{number}' where id={id}")
            self.conn.commit()
            return

    def get_current_date_ordinaries(self, district):
        with self.conn.cursor() as cursor:
            cursor.execute(
                f"SELECT * FROM sadidi_ordinary WHERE date_end >= CURRENT_DATE() AND sadik_id = ANY (SELECT id FROM sadidi_sadik WHERE district='{district}')")
            result = cursor.fetchall()
            return result

    def user_info(self, district):
        with self.conn.cursor() as cursor:
            cursor.execute(
                f"SELECT * FROM sadidi_ordinary WHERE date_end >= CURRENT_DATE() AND sadik_id = ANY (SELECT id FROM sadidi_sadik WHERE district='{district}')")
            result = cursor.fetchall()
            return result


if __name__ == "__main__":
    Database()
