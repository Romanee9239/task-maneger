function testMessage(){
  const cw = ChatWorkClient.factory({token: '4c87e81bedcfd3f62c2a0675d81e8af7'});
  const  body = 'テストメッセージ';
  cw.sendMessageToMyChat(body);
}