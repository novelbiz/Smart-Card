# -*- coding: utf-8 -*-
"""
# Copyright 2025 NOVELBIZ CO., LTD.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
"""

from smartcard.System import readers
from smartcard.CardType import AnyCardType
from smartcard.CardRequest import CardRequest
from smartcard.Exceptions import NoCardException
from smartcard.util import toHexString
from smartcard.scard import SCARD_PROTOCOL_T0, SCARD_PROTOCOL_T1, SCARD_SHARE_SHARED
import subprocess
import time


class IDCardReader:
    """
    Thai National ID Card Reader
    
    A simple command-line tool to read data from Thai National ID Cards
    using a smart card reader.
    
    Author: Novelbiz (https://novelbiz.co.th)
    License: MIT
    """
    
    def __init__(self):
        self.cardservice = None
        print("\n" + "="*70)
        print(" "*15 + "โปรแกรมอ่านบัตรประชาชนไทย - Novelbiz")
        print(" "*25 + "MIT License 2025")
        print("="*70)
        print("[เริ่มต้น] โปรแกรมพร้อมทำงาน")
        print("="*70 + "\n")

    # ------------------- Helper Functions -------------------
    def decode_text(self, data):
        """แปลง bytes เป็น text"""
        try:
            return bytes(data).decode('tis-620', errors='ignore').strip()
        except:
            return ''.join(chr(b) if b < 128 else '?' for b in data).strip()

    def send_apdu_with_get_response(self, connection, apdu):
        """ส่ง APDU command และจัดการ GET_RESPONSE"""
        response, sw1, sw2 = connection.transmit(apdu)
        if sw1 == 0x61:
            get_response = [0x00, 0xC0, 0x00, 0x00, sw2]
            response, sw1, sw2 = connection.transmit(get_response)
        return response, sw1, sw2

    def parse_thai_date(self, date_str):
        """แปลงวันที่จากรูปแบบ YYYYMMDD เป็นรูปแบบที่อ่านง่าย"""
        if date_str == '99999999':
            return "ตลอดชีพ", "LIFELONG"
            
        if len(date_str) == 8 and date_str.isdigit():
            try:
                year = date_str[0:4]
                month = date_str[4:6]
                day = date_str[6:8]
                
                thai_year = int(year)
                eng_year = thai_year - 543
                
                thai_months = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
                            'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม']
                
                eng_months = ['', 'January', 'February', 'March', 'April', 'May', 'June',
                            'July', 'August', 'September', 'October', 'November', 'December']
                
                month_int = int(month)
                if 1 <= month_int <= 12:
                    thai_date = f"{int(day)} {thai_months[month_int]} {thai_year}"
                    eng_date = f"{int(day)} {eng_months[month_int]} {eng_year}"
                    return thai_date, eng_date
            except Exception as e:
                print(f"[ผิดพลาด] แปลงวันที่ไม่สำเร็จ: {e}")
                return "ไม่ระบุ", "Not specified"
        else:
            return "ไม่ระบุ", "Not specified"

    def disconnect_card(self):
        """ตัดการเชื่อมต่อจากบัตร"""
        if self.cardservice:
            try:
                self.cardservice.connection.disconnect()
                print("[ตัดการเชื่อมต่อ] ตัดการเชื่อมต่อบัตรสำเร็จ")
            except Exception as e:
                print(f"[ผิดพลาด] ไม่สามารถตัดการเชื่อมต่อ: {e}")
            finally:
                self.cardservice = None

    # ------------------- Card Reader Functions -------------------
    def check_service_status(self):
        """ตรวจสอบสถานะ Smart Card Service"""
        try:
            result = subprocess.run(
                ["sc", "query", "SCardSvr"],
                capture_output=True, text=True, shell=True
            )
            return "RUNNING" in result.stdout
        except Exception as e:
            print(f"[ผิดพลาด] ตรวจสอบ Service ไม่สำเร็จ: {e}")
            return False

    def check_reader_status(self):
        """ตรวจสอบสถานะเครื่องอ่านบัตร"""
        if not self.check_service_status():
            print("[สถานะ] บริการ Smart Card ไม่ทำงาน")
            return False

        r = readers()
        if r:
            print(f"[สถานะ] พบเครื่องอ่านบัตร: {r[0]}")
            return True
        else:
            print("[สถานะ] ไม่พบเครื่องอ่านบัตร")
            return False

    # ------------------- Read ID Card -------------------
    def read_id_card(self):
        """อ่านข้อมูลจากบัตรประชาชน"""
        print("\n" + "="*70)
        print("[เริ่มอ่าน] กำลังอ่านข้อมูลบัตรประชาชน...")
        print("="*70)
        
        try:
            # ตัดการเชื่อมต่อเก่าก่อน (ถ้ามี)
            self.disconnect_card()
            
            # เชื่อมต่อกับบัตร
            cardtype = AnyCardType()
            cardrequest = CardRequest(timeout=5, cardType=cardtype)
            self.cardservice = cardrequest.waitforcard()
            
            print(f'\n[เชื่อมต่อ] เชื่อมต่อกับ: {self.cardservice.connection.getReader()}')
            
            try:
                self.cardservice.connection.connect(
                    protocol=SCARD_PROTOCOL_T0 | SCARD_PROTOCOL_T1,
                    mode=SCARD_SHARE_SHARED
                )
                atr = self.cardservice.connection.getATR()
                print(f'[ATR] {toHexString(atr)}')
            except Exception as e:
                print(f"[ผิดพลาด] ไม่สามารถเชื่อมต่อบัตรได้: {e}")
                self.disconnect_card()
                return False
       
            # เลือก Thai ID card applet
            SELECT = [0x00, 0xA4, 0x04, 0x00, 0x08]
            THAI_ID_CARD = [0xA0, 0x00, 0x00, 0x00, 0x54, 0x48, 0x00, 0x01]
            response, sw1, sw2 = self.cardservice.connection.transmit(SELECT + THAI_ID_CARD)
            
            # จัดการ GET RESPONSE สำหรับ SW 61
            if sw1 == 0x61:
                print(f"[ข้อมูล] กำลังดึงข้อมูลเพิ่มเติม... (SW: {sw1:02x} {sw2:02x})")
                get_response = [0x00, 0xC0, 0x00, 0x00, sw2]
                response, sw1, sw2 = self.cardservice.connection.transmit(get_response)
            
            if sw1 != 0x90:
                print(f"[ผิดพลาด] ไม่สามารถเลือก Applet ได้ SW: {sw1:02x} {sw2:02x}")
                self.disconnect_card()
                return False
            
            print(f"[สำเร็จ] เลือก Thai ID Applet สำเร็จ (SW: {sw1:02x} {sw2:02x})")

            # คำสั่งอ่านข้อมูลต่างๆ
            commands = {
                'cid': [0x80, 0xb0, 0x00, 0x04, 0x02, 0x00, 0x0d],
                'name_th': [0x80, 0xb0, 0x00, 0x11, 0x02, 0x00, 0x64],
                'name_en': [0x80, 0xb0, 0x00, 0x75, 0x02, 0x00, 0x64],
                'birth': [0x80, 0xb0, 0x00, 0xD9, 0x02, 0x00, 0x08],
                'gender': [0x80, 0xb0, 0x00, 0xE1, 0x02, 0x00, 0x01],
                'issuer': [0x80, 0xb0, 0x00, 0xF6, 0x02, 0x00, 0x64],
                'issue_date': [0x80, 0xb0, 0x01, 0x67, 0x02, 0x00, 0x08],
                'expire_date': [0x80, 0xb0, 0x01, 0x6F, 0x02, 0x00, 0x08],
                'address': [0x80, 0xb0, 0x15, 0x79, 0x02, 0x00, 0x64],
                'request_number': [0x80, 0xB0, 0x16, 0x19, 0x02, 0x00, 0x0E]
            }

            print("\n" + "-"*70)
            print(" "*25 + "ข้อมูลบัตรประชาชน")
            print("-"*70)

            # อ่านเลขบัตรประชาชน
            response, sw1, sw2 = self.send_apdu_with_get_response(
                self.cardservice.connection, commands['cid']
            )
            cid = self.decode_text(response)
            print(f"\n📌 เลขบัตรประชาชน: {cid}")

            # อ่านชื่อ-นามสกุล (ไทย)
            response, sw1, sw2 = self.send_apdu_with_get_response(
                self.cardservice.connection, commands['name_th']
            )
            name_th = self.decode_text(response).replace('#', ' ')
            print(f"👤 ชื่อ-นามสกุล (ไทย): {name_th}")

            # อ่านชื่อ-นามสกุล (อังกฤษ)
            response, sw1, sw2 = self.send_apdu_with_get_response(
                self.cardservice.connection, commands['name_en']
            )
            name_en = self.decode_text(response).replace('#', ' ')
            print(f"👤 ชื่อ-นามสกุล (อังกฤษ): {name_en}")

            # อ่านวันเกิด
            response, sw1, sw2 = self.send_apdu_with_get_response(
                self.cardservice.connection, commands['birth']
            )
            birth_raw = self.decode_text(response)
            birth_th, birth_en = self.parse_thai_date(birth_raw)
            print(f"🎂 วันเกิด (ไทย): {birth_th}")
            print(f"🎂 วันเกิด (อังกฤษ): {birth_en}")

            # อ่านเพศ
            response, sw1, sw2 = self.send_apdu_with_get_response(
                self.cardservice.connection, commands['gender']
            )
            gender_code = self.decode_text(response)
            gender = "ชาย" if gender_code == "1" else "หญิง" if gender_code == "2" else gender_code
            print(f"⚧ เพศ: {gender}")

            # อ่านวันที่ออกบัตร
            response, sw1, sw2 = self.send_apdu_with_get_response(
                self.cardservice.connection, commands['issue_date']
            )
            issue_raw = self.decode_text(response)
            issue_th, issue_en = self.parse_thai_date(issue_raw)
            print(f"\n📅 วันที่ออกบัตร (ไทย): {issue_th}")
            print(f"📅 วันที่ออกบัตร (อังกฤษ): {issue_en}")

            # อ่านวันหมดอายุ
            response, sw1, sw2 = self.send_apdu_with_get_response(
                self.cardservice.connection, commands['expire_date']
            )
            expire_raw = self.decode_text(response)
            expire_th, expire_en = self.parse_thai_date(expire_raw)
            print(f"📅 วันที่หมดอายุ (ไทย): {expire_th}")
            print(f"📅 วันที่หมดอายุ (อังกฤษ): {expire_en}")
            
            # อ่านสถานที่ออกบัตร
            response, sw1, sw2 = self.send_apdu_with_get_response(
                self.cardservice.connection, commands['issuer']
            )
            issuer_name = self.decode_text(response)
            print(f"🏢 สถานที่ออกบัตร: {issuer_name}")
            
            # อ่านที่อยู่
            response, sw1, sw2 = self.send_apdu_with_get_response(
                self.cardservice.connection, commands['address']
            )
            address = self.decode_text(response).replace('#', ' ')
            print(f"\n🏠 ที่อยู่: {address}")
            
            # อ่านหมายเลขคำขอ
            response, sw1, sw2 = self.send_apdu_with_get_response(
                self.cardservice.connection, commands['request_number']
            )
            request_number = self.decode_text(response)
            print(f"🔢 หมายเลขคำขอ: {request_number}")

            # ตัดการเชื่อมต่อหลังจากอ่านเสร็จ
            self.disconnect_card()
            
            print("\n" + "="*70)
            print("[สำเร็จ] ✅ อ่านข้อมูลบัตรประชาชนเรียบร้อยแล้ว")
            print("="*70 + "\n")
            return True
            
        except NoCardException:
            print("\n[ผิดพลาด] ไม่พบบัตรประชาชน กรุณาใส่บัตรแล้วลองใหม่")
            self.disconnect_card()
            return False
        except Exception as e:
            print(f"\n[ผิดพลาด] เกิดข้อผิดพลาดในการอ่านบัตร: {e}")
            self.disconnect_card()
            return False

    # ------------------- Main Loop -------------------
    def run(self):
        """รันโปรแกรมหลัก"""
        # ตรวจสอบเครื่องอ่านบัตรก่อน
        if not self.check_reader_status():
            print("\n[ผิดพลาด] ไม่พบเครื่องอ่านบัตร กรุณาเชื่อมต่อเครื่องอ่านบัตรแล้วรันโปรแกรมใหม่")
            return
        
        print("\n[พร้อม] ใส่บัตรประชาชนเพื่ออ่านข้อมูล...")
        
        try:
            # อ่านบัตรครั้งเดียว
            if self.read_id_card():
                print("\n[เสร็จสิ้น] อ่านข้อมูลเสร็จสมบูรณ์")
            else:
                print("\n[ผิดพลาด] ไม่สามารถอ่านข้อมูลได้")
                
        except KeyboardInterrupt:
            print("\n\n[ยกเลิก] ยกเลิกการอ่านข้อมูลโดยผู้ใช้")
        except Exception as e:
            print(f"\n[ผิดพลาด] เกิดข้อผิดพลาด: {e}")


# ------------------- Main -------------------
if __name__ == "__main__":
    print("\n")
    print("*"*70)
    print("*" + " "*68 + "*")
    print("*" + " "*10 + "โปรแกรมอ่านบัตรประชาชนไทย (CLI Version)" + " "*19 + "*")
    print("*" + " "*25 + "by Novelbiz" + " "*34 + "*")
    print("*" + " "*22 + "MIT License 2025" + " "*33 + "*")
    print("*" + " "*68 + "*")
    print("*"*70)
    
    try:
        reader = IDCardReader()
        reader.run()
    except KeyboardInterrupt:
        print("\n\n[ออก] ปิดโปรแกรมโดยผู้ใช้")
    except Exception as e:
        print(f"\n[ผิดพลาดร้ายแรง] {e}")
    finally:
        print("\n" + "="*70)
        print("ขอบคุณที่ใช้โปรแกรมจาก Novelbiz")
        print("Website: https://novelbiz.co.th")
        print("="*70 + "\n")
