import io, tempfile, unittest
import app as target
class Fake:
    def reply(self,messages):
        assert "Календарь не подключён" in messages[0]["content"]
        return "Тестовый ответ"
class TestBot(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.NamedTemporaryFile(suffix='.db',delete=False);target.DB=target.Path(self.tmp.name);target.gigachat=Fake();target.app.config['TESTING']=True;self.c=target.app.test_client()
    def test_upload_and_chat(self):
        h={'X-Admin-Password':'admin123'}
        r=self.c.post('/api/admin/upload',headers=h,data={'files':(io.BytesIO('Эксперт проводит консультации.'.encode()),'base.txt')},content_type='multipart/form-data');self.assertEqual(r.status_code,200)
        r=self.c.post('/api/chat',json={'message':'Как проходит консультация?'});self.assertEqual(r.json['answer'],'Тестовый ответ')
if __name__=='__main__': unittest.main()
