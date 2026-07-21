$ids = @('bde46773-e0fa-4273-9d6e-2a74f0737d1c','738e8e52-466e-4352-8a8a-407085938625','486041cf-801f-472b-97d2-393a4ec5f28e','ed47615e-ec0f-4796-99b6-562f54e8f851','9cb8f7e7-72d1-447a-895c-ad9227d36a4c','0fe5a348-fb97-4427-aacc-c602902d0b98','0e7bc7a9-5d86-49b8-990b-7a28d142a5f8','c7802546-5527-41b1-b2ec-54861f91914a','62fe2681-bc5a-4b0c-8f7d-f9228bf6a507','940737e4-3485-4e20-8305-287ba862ba1e','5a6e105a-09a0-47a9-99f6-d3d41cb90c78','f810ed71-6d50-4e7f-80c6-e81c5fbb12ac','163e432c-c300-4e96-97d7-0c49826faa16')
Connect-MgGraph -Scopes "Group.Read.All"
$ids | ForEach-Object {
  try { (Get-MgGroup -GroupId $_).DisplayName } catch { "$_ : Not found or access denied" }
}
